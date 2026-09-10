// Use the parser shipped with the installed Microsoft Power Query extension, without changing it.
const fs = require('node:fs');
const vm = require('node:vm');
const [bundlePath, sourcePath] = process.argv.slice(2);
let bundle = fs.readFileSync(bundlePath, 'utf8');
const entry = 'var l=t(1422);module.exports=l';
if (!bundle.includes(entry)) throw new Error('Unrecognized installed parser bundle entry point.');
bundle = bundle.replace(entry, 'module.exports=t(9833)');
const sandbox = { module: {exports:{}}, require, console, process, Buffer, setTimeout, clearTimeout, setImmediate, clearImmediate };
vm.runInNewContext(bundle, sandbox, {filename: bundlePath});
const pq = sandbox.module.exports;
(async () => {
  const result = await pq.TaskUtils.tryLexParse(pq.DefaultSettings, fs.readFileSync(sourcePath, 'utf8'));
  if (pq.TaskUtils.isError(result)) {
    console.error(JSON.stringify(result.error.innerError || {message:result.error.message}, null, 2)); process.exitCode = 1;
  } else {
    console.log(JSON.stringify({source:sourcePath, syntax:'passed', stage:result.stage}));
  }
})();
