import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import test from 'node:test';
import { build } from 'esbuild';
import { compareTableText, compareFilterValues } from '../src/app/components/sort-utils';

test('period changes reuse source rows in the actual calculation hook; source edits replace and release that worker', async () => {
  const bundle = await build({ stdin: { resolveDir: process.cwd(), contents: `
    import {useTimesheetCalculations} from './src/app/hooks/useTimesheetCalculations';
    import {begin,commit,unmount} from 'hook-runtime';
    export {unmount};
    const empty=[];
    export function render(rows,from,to){begin();const value=useTimesheetCalculations(rows,empty,empty,empty,from,to);commit();return value;}
  ` }, bundle: true, write: false, platform: 'node', format: 'cjs', packages: 'external',
    plugins: [{ name: 'filter-session-boundaries', setup(builder) {
      builder.onResolve({ filter: /^react$|^hook-runtime$/ }, () => ({ path: 'react', namespace: 'fixture' }));
      builder.onResolve({ filter: /AppDataContext$|ui-settings$|timesheet\.worker\?worker$/ }, args => ({ path: args.path, namespace: 'fixture' }));
      builder.onLoad({ filter: /.*/, namespace: 'fixture' }, ({path}) => ({ contents: path === 'react' ? `
        const slots=[];let cursor=0;const effects=[];
        const same=(a,b)=>a&&b&&a.length===b.length&&a.every((v,i)=>Object.is(v,b[i]));
        export function begin(){cursor=0;}
        export function commit(){for(const run of effects.splice(0))run();}
        export function unmount(){for(const slot of slots)slot.cleanup?.();}
        export function useRef(value){const i=cursor++;return slots[i]||=( {current:value} );}
        export function useState(initial){const i=cursor++;const slot=slots[i]||=( {value:typeof initial==='function'?initial():initial} );return [slot.value,value=>{slot.value=typeof value==='function'?value(slot.value):value;}];}
        export function useMemo(fn,deps){const i=cursor++;if(!slots[i]||!same(slots[i].deps,deps))slots[i]={value:fn(),deps};return slots[i].value;}
        export const useCallback=(fn,deps)=>useMemo(()=>fn,deps);
        export function useEffect(fn,deps){const i=cursor++;const slot=slots[i]||=( {} );if(!same(slot.deps,deps))effects.push(()=>{slot.cleanup?.();slot.deps=deps;slot.cleanup=fn();});}
        export const useSyncExternalStore=(_subscribe,get)=>get();
      ` : path.endsWith('AppDataContext') ? `
        const appData={Q_CheckTAs:[],Timesheet_InputList:[{id:'source',l07:'MKT LOCAL NORTH'}]};
        export const useAppData=()=>({appData});
      ` : path.endsWith('ui-settings') ? 'export const useUiSettings=()=>({defaultAuditYear:2026});'
        : 'export default FakeWorker;' }));
    } }],
  });
  type Message = {type: string; requestId: string; rows?: unknown[]; params?: {fromDateStr: string; toDateStr: string}};
  const workers: FakeWorker[] = [];
  class FakeWorker {
    onmessage: ((event: {data: unknown}) => void) | null = null;
    onerror: ((event: unknown) => void) | null = null;
    onmessageerror: (() => void) | null = null;
    messages: Message[] = [];
    terminated = false;
    constructor(){workers.push(this);}
    postMessage(message: Message){
      this.messages.push(message);
      if(message.type==='timesheet-input-complete'||message.type==='timesheet-recalculate')queueMicrotask(()=>this.onmessage?.({data:{requestId:message.requestId,processedRosterData:[{id:message.requestId}],employeeSummary:[],centerSummary:[]}}));
    }
    terminate(){this.terminated=true;}
  }
  const module={exports:{} as {render:(rows: unknown[],from: string,to: string)=>{error?: string; retry:()=>void};unmount:()=>void}};
  new Function('require','module','exports','FakeWorker',bundle.outputFiles[0].text)(createRequire(import.meta.url),module,module.exports,FakeWorker);
  const waitFor=async (condition:()=>boolean)=>{const end=Date.now()+5000;while(!condition()){if(Date.now()>end)throw new Error('Filter worker did not finish');await new Promise(resolve=>setTimeout(resolve,10));}};
  const rows=Array.from({length:4500},(_,i)=>({id:i}));
  try {
    module.exports.render(rows,'2026-08-21','2026-09-20');
    await waitFor(()=>workers[0]?.messages.some(message=>message.type==='timesheet-input-complete'));
    const chunks=workers[0].messages.filter(message=>message.type==='timesheet-input-chunk');
    assert.equal(chunks.reduce((sum,message)=>sum+message.rows!.length,0),4500);
    assert.ok(chunks.every(message=>message.rows!.length<=2000));
    module.exports.render(rows,'2026-09-21','2026-10-20');
    module.exports.render(rows,'2026-10-21','2026-11-20');
    await waitFor(()=>workers[0].messages.some(message=>message.type==='timesheet-recalculate'));
    assert.equal(workers.length,1);
    assert.equal(workers[0].messages.filter(message=>message.type==='timesheet-input-chunk').length,chunks.length);
    const query=workers[0].messages.find(message=>message.type==='timesheet-recalculate')!;
    assert.equal(query.params!.toDateStr,'2026-11-20');
    const edited=[...rows,{id:'new edit'}];
    module.exports.render(edited,'2026-10-21','2026-11-20');
    await waitFor(()=>workers[1]?.messages.some(message=>message.type==='timesheet-input-complete'));
    assert.equal(workers[0].terminated,true);
    assert.equal(workers[1].messages.filter(message=>message.type==='timesheet-input-chunk').reduce((sum,message)=>sum+message.rows!.length,0),4501);
    workers[1].onmessageerror!();
    assert.equal(workers[1].terminated,true);
    const failed=module.exports.render(edited,'2026-10-21','2026-11-20');
    assert.match(failed.error!,/Không thể nhận kết quả/);
    failed.retry();
    module.exports.render(edited,'2026-10-21','2026-11-20');
    await waitFor(()=>workers[2]?.messages.some(message=>message.type==='timesheet-input-complete'));
    assert.equal(workers.length,3,'retry replaces a failed worker without losing the source edits');
  } finally {module.exports.unmount();}
  assert.ok(workers.every(worker=>worker.terminated));
});

test('reused text collators preserve Vietnamese, accents, numeric names and filter ordering', () => {
  const values=['An Hưng','AN HUNG','Đặng','DANG','KDG 2','KDG 10','00123','123','HN2','HN10',''];
  for(const left of values)for(const right of values){
    assert.equal(Math.sign(compareTableText(left,right)),Math.sign(left.localeCompare(right,'vi',{numeric:true,sensitivity:'base'})));
    assert.equal(Math.sign(compareFilterValues(left,right)),Math.sign(left.localeCompare(right,undefined,{numeric:true})));
  }
});
