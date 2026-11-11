const fs=require('fs'),path=require('path'),cp=require('child_process');
const root=__dirname;
const cases=[
 ['inline',String.raw`\(e_q=q_r-\hat q\)`,true],
 ['doubled',String.raw`\\[ q_r(t)=\\sum_i A_i v_1(t-t_i) \\]`,true],
 ['code','```latex\n\\tau=\\frac{u_0-z_3}{b_0}\n```',true],
 ['matrix',String.raw`\begin{bmatrix}q_r\\\dot q_r\end{bmatrix}`,true],
 ['invalid',String.raw`\frac{a}{`,false],
];
for(const [name,latex,ok] of cases){
 const base=path.join(root,'output','test_'+name),req=path.join(root,'temp','test_'+name+'.json');
 fs.writeFileSync(req,JSON.stringify({latex,fontPt:14,outputBase:base}));
 const result=cp.spawnSync(path.join(root,'vendor/node/node.exe'),[path.join(root,'render.cjs'),req],{encoding:'utf8',timeout:20000});
 if((result.status===0)!==ok)throw Error(name+': '+result.stderr);
 if(ok){const svg=fs.readFileSync(base+'.svg','utf8');if(!svg.includes('<path')||svg.includes('<script'))throw Error('Invalid vector '+name);}
 process.stdout.write('PASS '+name+'\n');
}
