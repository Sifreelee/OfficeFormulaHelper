const fs=require('fs'),path=require('path'),cp=require('child_process');
const root=__dirname;
const cases=[
 ['sub',String.raw`v_1`,['msub']],
 ['sup',String.raw`x^2`,['msup']],
 ['both',String.raw`q_{d,i}^{(2)}`,['msubsup']],
 ['dot',String.raw`\dot q_r`,['mover','msub']],
 ['sum',String.raw`\sum_{i=1}^{n} x_i^2`,['munderover','msubsup']]
];
for(const [name,latex,tags] of cases){
 const base=path.join(root,'output','scripts_'+name),req=path.join(root,'temp','scripts_'+name+'.json');
 fs.writeFileSync(req,JSON.stringify({latex,fontPt:18,outputBase:base}));
 const result=cp.spawnSync(path.join(root,'vendor/node/node.exe'),[path.join(root,'render.cjs'),req],{encoding:'utf8',timeout:20000});
 if(result.status!==0)throw Error(result.stderr);
 const mathml=fs.readFileSync(base+'.mml','utf8');
 for(const tag of tags)if(!new RegExp('<'+tag+'(?:\\s|>)').test(mathml))throw Error(name+' lost '+tag);
 process.stdout.write('PASS script structure '+name+'\n');
}
