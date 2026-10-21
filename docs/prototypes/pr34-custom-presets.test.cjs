'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const htmlPath = path.join(__dirname, 'PR34_CUSTOM_PRESETS.html');
const html = fs.readFileSync(htmlPath, 'utf8');
const match = html.match(/data-srcdoc="([\s\S]*?)"/);
assert(match, 'standalone wrapper contains the source document');
const entities = {'amp':'&', 'lt':'<', 'gt':'>', 'quot':'"', '#x27':"'", '#39':"'"};
const source = match[1].replace(/&(amp|lt|gt|quot|#x27|#39);/g, (_, key) => entities[key]);
const scripts = [...source.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)];
const original = scripts.map(m => m[1]).find(s => s.includes("getElementById('custom-preset-fem')"));
assert(original, 'prototype script found');
new vm.Script(original);
const marker = 'restore(window.openai?.widgetState);resetDrafts();render();';
assert.equal(original.split(marker).length, 2, 'one test instrumentation point');
const instrumented = original.replace(marker, `
globalThis.testAPI = {
  buildPackage, validatePackage, registerJSON, packageKey, saveCustom, loadCustom,
  applyPreset, mappingErrors, validate, persist, restore, canonical,
  state: () => ({docs,library,example,selected,opId,pending,sequence,catalog}),
  select: key => {example=key;selected=stages()[0]?.id||'';opId=stage()?.ops[0]?.id||'';resetDrafts();},
  pending: value => {pending=value;},
};
${marker}`);

function environment(hostThrows = false) {
  const nodes = new Map();
  const node = id => {
    if (!nodes.has(id)) nodes.set(id, {innerHTML:'',textContent:'',value:'',hidden:false,
      classList:{toggle(){}},addEventListener(){}});
    return nodes.get(id);
  };
  const root = node('root');
  root.querySelector = node;
  root.contains = () => true;
  let saved;
  const context = {document:{getElementById:()=>root},window:{addEventListener(){},openai:{
    setWidgetState(value) {
      if (hostThrows) throw new Error('optional host state unavailable');
      saved = JSON.parse(JSON.stringify(value));
      return Promise.resolve();
    },
  }}};
  vm.runInNewContext(instrumented, context);
  return {api:context.testAPI, saved:()=>saved};
}
const copy = value => JSON.parse(JSON.stringify(value));
const meta = (scope='FLOW_ONLY', extra={}) => ({presetId:'U_TEST',revision:1,name:'自作試験',description:'入力構成の保存・読込試験',scope,...extra});
const protectedState = api => api.canonical({docs:api.state().docs,library:api.state().library,sequence:api.state().sequence});
const currentDoc = api => api.state().docs[api.state().example];
const flowShape = doc => doc.stages.map(st => ({name:st.name,
  parent:st.parent==='INITIAL'?-1:doc.stages.findIndex(s=>s.id===st.parent),
  ops:st.ops.map(op=>({kind:op.kind,targetType:op.target?doc.sets.find(s=>s.id===op.target)?.type:null,
    p:{...op.p,segment:op.p.segment?st.ops.findIndex(o=>o.id===op.p.segment):''}}))}));
let passed = 0;
function check(name, fn) {fn();passed++;console.log('PASS '+name);}

check('FLOW_ONLY strips numeric inputs and preserves internal references', () => {
  const {api} = environment();
  api.applyPreset('slope_seismic');
  const doc = currentDoc(api);
  doc.stages[1].ops[0].p.value = '0.15';
  const packet = api.buildPackage(meta());
  api.validatePackage(packet);
  assert.equal(packet.stages[1].ops[0].p.value, '');
  assert.equal(packet.slots.length, 1);
  assert.equal(packet.slots[0].active, null);
  assert(packet.stages[1].ops.some(o=>o.kind==='srm'));
  assert(!Object.hasOwn(packet, 'results'));
  assert.deepEqual(copy(packet.externalBindings), []);
});

check('MODEL_INPUTS round-trip preserves target definitions and entered values', () => {
  const {api} = environment();
  api.applyPreset('slope_seismic');
  const originalDoc = currentDoc(api);
  originalDoc.stages[1].ops[0].p.value = '-0.12';
  const expected = JSON.stringify(flowShape(originalDoc));
  const priorDocs = api.canonical(api.state().docs);
  const packet = api.buildPackage(meta('MODEL_INPUTS'));
  const key = api.registerJSON(JSON.stringify(packet));
  const oldIds = new Set(Object.values(api.state().docs).flatMap(d=>[
    ...d.sets.map(s=>s.id),...d.stages.flatMap(s=>[s.id,...s.ops.map(o=>o.id)])]));
  api.pending({type:'custom-load',key,mode:'FULL_NEW',map:{},error:''});
  api.loadCustom();
  const loaded = currentDoc(api);
  assert.equal(JSON.stringify(flowShape(loaded)), expected);
  for (const id of [...loaded.sets.map(t=>t.id),...loaded.stages.flatMap(s=>[s.id,...s.ops.map(o=>o.id)])]) assert(!oldIds.has(id));
  const newKey = api.state().example;
  const remaining = {...api.state().docs};delete remaining[newKey];
  assert.equal(api.canonical(remaining), priorDocs);
  assert.equal(api.canonical(api.state().library[key]), api.canonical(packet));
  const resaved = api.buildPackage(meta('MODEL_INPUTS',{presetId:'U_ROUNDTRIP'}));
  api.validatePackage(resaved);
  assert.equal(resaved.stages[1].ops[0].p.value, '-0.12');
});

check('mapped loading requires explicit compatible mapping and preserves source input', () => {
  const {api} = environment();
  const packet = api.buildPackage(meta());
  const key = api.registerJSON(JSON.stringify(packet));
  const before = protectedState(api);
  api.pending({type:'custom-load',key,mode:'MAPPED_TO_CURRENT_INPUT',map:{},error:''});
  api.loadCustom();
  assert.match(api.state().pending.error, /未指定/);
  assert.equal(protectedState(api), before);
  api.pending({type:'custom-load',key,mode:'MAPPED_TO_CURRENT_INPUT',map:{G:'D'},error:''});
  api.loadCustom();
  assert.match(api.state().pending.error, /型/);
  assert.equal(protectedState(api), before);
  api.pending({type:'custom-load',key,mode:'MAPPED_TO_CURRENT_INPUT',map:{G:'G'},error:''});
  api.loadCustom();
  const loaded = currentDoc(api);
  assert.equal(loaded.mappingMode, 'MAPPED_TO_CURRENT_INPUT');
  assert.equal(loaded.sets.length, 3, 'whole current model context is cloned, as previewed');
  assert.equal(loaded.mappings[0].sourceTarget, 'G');
  assert(loaded.sets.some(t=>t.id===loaded.mappings[0].newTarget));
  assert.notEqual(loaded.mappings[0].newTarget, 'G');
});

check('saved revisions are immutable; name and identity collisions reject without mutation', () => {
  const {api} = environment();
  api.pending({type:'custom-save',base:'new',name:'自作A',description:'初版',scope:'FLOW_ONLY',error:''});
  api.saveCustom();
  const first = Object.values(api.state().library)[0];
  const originalFirst = api.canonical(first);
  api.pending({type:'custom-save',base:first.presetId,name:'自作A',description:'改訂',scope:'MODEL_INPUTS',error:''});
  api.saveCustom();
  assert.equal(Object.keys(api.state().library).length, 2);
  assert.equal(api.canonical(api.state().library[api.packageKey(first)]), originalFirst);
  api.pending({type:'custom-save',base:'new',name:'自作B',description:'',scope:'FLOW_ONLY',error:''});
  api.saveCustom();
  const before = protectedState(api);
  api.pending({type:'custom-save',base:first.presetId,name:'自作B',description:'',scope:'FLOW_ONLY',error:''});
  api.saveCustom();
  assert.match(api.state().pending.error, /同じ名称/);
  assert.equal(protectedState(api), before);
  const changed = copy(first);changed.description='same identity, different content';
  assert.throws(()=>api.registerJSON(JSON.stringify(changed)), /同じID/);
  assert.equal(protectedState(api), before);
  api.registerJSON(JSON.stringify(first));
  assert.equal(protectedState(api), before, 'identical import is idempotent');
});

check('malformed data, unknown schema, duplicate IDs, bad references and units are nonmutating', () => {
  const {api} = environment();
  const packet = api.buildPackage(meta('MODEL_INPUTS'));
  const before = protectedState(api);
  const mutations = [
    p=>{p.schemaVersion=99;}, p=>{p.dictionaryVersion=99;},
    p=>{p.units='mm-N-s';}, p=>{p.stages[1].id=p.stages[0].id;},
    p=>{p.stages[1].ops[1].p.segment='O_MISSING';},
    p=>{p.stages[0].ops[0].kind='RUN_MACRO';},
    p=>{p.stages[0].ops[0].p.value='(()=>{})()';},
    p=>{p.results={};}, p=>{p.externalBindings=['committed-state'];},
    p=>{p.slots[0].type='beam';}, p=>{p.stages[0].parent=p.stages[1].id;},
    p=>{p.description='a'.repeat(4001);},
  ];
  for (const change of mutations) {
    const bad=copy(packet);change(bad);
    assert.throws(()=>api.registerJSON(JSON.stringify(bad)));
    assert.equal(protectedState(api), before);
  }
  assert.throws(()=>api.registerJSON('{not JSON'));
  assert.throws(()=>api.registerJSON(' '.repeat(131073)));
  assert.equal(protectedState(api), before);
});

check('all eight built-ins remain read-only while custom copies can be saved', () => {
  const {api} = environment();
  const original=api.canonical(api.state().catalog);
  for (const key of Object.keys(api.state().catalog)) {
    api.applyPreset(key);
    const packet=api.buildPackage(meta('MODEL_INPUTS',{presetId:'U_'+key,name:key}));
    api.validatePackage(packet);
    api.registerJSON(JSON.stringify(packet));
  }
  assert.equal(api.canonical(api.state().catalog), original);
});

check('host state caching is optional and JSON remains the portable save representation', () => {
  const {api} = environment(true);
  api.pending({type:'custom-save',base:'new',name:'一覧保存',description:'',scope:'FLOW_ONLY',error:''});
  api.saveCustom();
  assert.equal(Object.keys(api.state().library).length, 1);
  const packet=Object.values(api.state().library)[0];
  const {api:other} = environment();
  other.registerJSON(JSON.stringify(packet));
  assert.equal(other.canonical(Object.values(other.state().library)[0]), other.canonical(packet));
});

check('bounded host snapshot restores library and input together', () => {
  const env=environment(), {api}=env;
  api.pending({type:'custom-save',base:'new',name:'復元確認',description:'',scope:'FLOW_ONLY',error:''});
  api.saveCustom();api.persist();
  const saved=env.saved();assert(saved.privateContent);
  assert(Buffer.byteLength(JSON.stringify(saved))<16384);
  const {api:restored}=environment();restored.restore(saved);
  assert.equal(restored.canonical(restored.state().docs), api.canonical(api.state().docs));
  assert.equal(restored.canonical(restored.state().library), api.canonical(api.state().library));
});

const fixturePath=path.join(__dirname,'pr34-custom-preset-example.json');
if(process.argv.includes('--write-fixture')) {
  const {api}=environment();
  fs.writeFileSync(fixturePath, JSON.stringify(api.buildPackage(meta('FLOW_ONLY',{
    presetId:'U_DEMO',name:'常時斜面の照査工程（自作例）',description:'工程構成のみ。読込時に地盤slotを明示対応付けし、数値と材料条件を再確認。',
  })),null,2)+'\n');
}
check('published JSON example can be imported as a custom preset',()=>{
  const {api}=environment();const raw=fs.readFileSync(fixturePath,'utf8');
  const key=api.registerJSON(raw);assert.equal(key,'U_DEMO@1');
});
console.log(`${passed} checks passed; no Excel or numerical solver execution.`);
