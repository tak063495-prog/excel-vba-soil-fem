// Optional symmetric sparse direct solver. Wrapper: Apache-2.0.
// Eigen 3.4.0 headers: MPL-2.0/BSD; see EIGEN_NOTICE.txt.
#define EIGEN_MPL2_ONLY
#include <Eigen/SparseCholesky>
#include <chrono>
#include <cmath>
#include <cstdint>
#include <memory>
#include <mutex>
#include <unordered_map>
#include <vector>

using Sparse = Eigen::SparseMatrix<double>;
using Clock = std::chrono::steady_clock;
struct BoundaryTerm { int row, fixed; double value; };
struct Factor {
    Sparse matrix;
    Eigen::SimplicialLDLT<Sparse> solver;
    Eigen::VectorXd scale;
    std::vector<int> fixed;
    std::vector<BoundaryTerm> boundary;
};
static std::mutex registryMutex;
static std::unordered_map<int64_t, std::unique_ptr<Factor>> registry;
static int64_t nextToken = 1;
#define API extern "C" __declspec(dllexport) int __stdcall

API sfs_abi() noexcept { return 0x10000; }

// VBA upper band a(row,offset), column-major. Exact Dirichlet elimination;
// no penalty diagonal, diagonal shift, entry dropping, or force-law changes.
API sfs_factor_band(int n, int bw, int ld, const double* band, const int* locked,
                    double limitBytes, int64_t* token, double* metrics) noexcept {
    if (!token || !metrics) return 100;
    *token = 0;
    if (n < 1 || n > 2000000 || bw < 0 || bw >= n || ld < n || !band || !locked ||
        !std::isfinite(limitBytes) || limitBytes <= 0) return 100;
    try {
        auto start = Clock::now();
        auto f = std::make_unique<Factor>();
        f->fixed.assign(locked, locked + n);
        f->scale.resize(n);
        std::vector<Eigen::Triplet<double>> entries;
        entries.reserve(static_cast<size_t>(n) * std::min(bw + 1, 40));
        for (int i = 0; i < n; ++i) {
            double d = locked[i] ? 1.0 : band[i];
            if (!std::isfinite(d) || d <= 0) return 2;
            f->scale[i] = 1.0 / std::sqrt(d);
            entries.emplace_back(i, i, d);
        }
        for (int j = 1; j <= bw; ++j) {
            const double* column = band + static_cast<int64_t>(j) * ld;
            for (int i = 0; i + j < n; ++i) {
                double v = column[i];
                if (!std::isfinite(v)) return 2;
                if (v == 0) continue;
                int k = i + j;
                if (!locked[i] && !locked[k]) {
                    entries.emplace_back(i, k, v);
                    entries.emplace_back(k, i, v);
                } else if (!locked[i]) f->boundary.push_back({i, k, v});
                else if (!locked[k]) f->boundary.push_back({k, i, v});
            }
        }
        f->matrix.resize(n, n);
        f->matrix.setFromTriplets(entries.begin(), entries.end());
        entries.clear(); entries.shrink_to_fit();
        Sparse scaled = f->matrix;
        for (int j = 0; j < n; ++j)
            for (Sparse::InnerIterator it(scaled, j); it; ++it)
                it.valueRef() *= f->scale[it.row()] * f->scale[it.col()];
        f->solver.compute(scaled);
        if (f->solver.info() != Eigen::Success) return 3;
        auto d = f->solver.vectorD();
        if (!d.allFinite() || d.minCoeff() <= 1e-14) return 3;
        double lnnz = static_cast<double>(f->solver.matrixL().nestedExpression().nonZeros());
        double bytes = (2.0 * f->matrix.nonZeros() + lnnz) * (sizeof(double) + sizeof(int)) +
                       static_cast<double>(n) * 128 + f->boundary.size() * sizeof(BoundaryTerm);
        if (bytes > limitBytes) return 7;
        metrics[0] = d.minCoeff(); metrics[1] = d.maxCoeff();
        metrics[2] = lnnz; metrics[3] = static_cast<double>(f->matrix.nonZeros());
        metrics[4] = std::chrono::duration<double>(Clock::now() - start).count(); metrics[5] = bytes;
        std::lock_guard<std::mutex> guard(registryMutex);
        if (nextToken <= 0 || nextToken == INT64_MAX) return 100;
        *token = nextToken++;
        registry.emplace(*token, std::move(f));
        return 0;
    } catch (const std::bad_alloc&) { return 7; }
      catch (...) { return 100; }
}

API sfs_solve(int64_t token, int n, const double* rhs, const double* prescribed,
              double tolerance, double* result, double* metrics) noexcept {
    if (!rhs || !prescribed || !result || !metrics || !std::isfinite(tolerance) || tolerance <= 0) return 100;
    try {
        std::lock_guard<std::mutex> guard(registryMutex);
        auto found = registry.find(token);
        if (found == registry.end() || found->second->matrix.rows() != n) return 4;
        auto start = Clock::now(); auto& f = *found->second;
        Eigen::VectorXd b(n);
        for (int i = 0; i < n; ++i) {
            if (!std::isfinite(rhs[i]) || !std::isfinite(prescribed[i])) return 2;
            b[i] = f.fixed[i] ? prescribed[i] : rhs[i];
        }
        for (auto& term : f.boundary) b[term.row] -= term.value * prescribed[term.fixed];
        if (!b.allFinite()) return 2;
        Eigen::VectorXd x = f.scale.cwiseProduct(f.solver.solve(f.scale.cwiseProduct(b)));
        if (f.solver.info() != Eigen::Success || !x.allFinite()) return 5;
        double norm = b.norm(), relative = 0;
        // Iterative refinement checks the actual eliminated, unscaled matrix.
        for (int step = 0; step <= 4; ++step) {
            Eigen::VectorXd r = b - f.matrix * x;
            relative = r.norm() / (norm > 0 ? norm : 1.0);
            if (!std::isfinite(relative)) return 5;
            if (relative <= tolerance || step == 4) break;
            x += f.scale.cwiseProduct(f.solver.solve(f.scale.cwiseProduct(r)));
            if (f.solver.info() != Eigen::Success || !x.allFinite()) return 5;
        }
        metrics[6] = relative;
        metrics[7] = std::chrono::duration<double>(Clock::now() - start).count();
        if (relative > tolerance) return 5;
        Eigen::Map<Eigen::VectorXd>(result, n) = x;
        return 0;
    } catch (const std::bad_alloc&) { return 7; }
      catch (...) { return 100; }
}

API sfs_release(int64_t token) noexcept {
    try { std::lock_guard<std::mutex> guard(registryMutex); registry.erase(token); return 0; }
    catch (...) { return 100; }
}
