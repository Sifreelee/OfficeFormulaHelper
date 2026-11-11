const fs=require('fs'),path=require('path');
const runtime='C:/Users/lenovo/.cache/codex-runtimes/codex-primary-runtime/dependencies/node';
fs.cpSync(path.join(runtime,'bin'),path.join(__dirname,'vendor/node'),{recursive:true,dereference:true});
for(const name of ['sharp','@img/colour','detect-libc','semver','@img/sharp-win32-x64']){
 const source=fs.realpathSync(path.join(runtime,'node_modules',name));
 fs.cpSync(source,path.join(__dirname,'vendor/node_modules',name),{recursive:true,dereference:true});
}
fs.writeFileSync(path.join(__dirname,'runtime.json'),JSON.stringify({node:'vendor/node/node.exe',modules:'vendor/node_modules'},null,2));
