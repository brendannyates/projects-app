/* Demo Library navigation. Loads only with ?showcase=1; never alters workbook data. */
(() => {
  if (!new URLSearchParams(location.search).has('showcase') || document.getElementById('demoHubNav')) return;
  const base = 'https://brendannyates.github.io';
  const items = [
    ['TMS Tracker', base+'/tms-tracker/demo-v2-original.html'],
    ['ESK Tracker', base+'/esk-tracker/taskpane.html?demo=1&showcase=1'],
    ['IR Portal', base+'/IR-Portal/taskpane.html?demo=1&showcase=1'],
    ['IVK Tracker', base+'/ivk-tracker/taskpane.html?showcase=1'],
    ['Sohar Eligibility', base+'/sohar-eligibility-addin/dashboard.html?fees=1&showcase=1#benefits'],
    ['MHS Clinic App', base+'/mhs-clinic-app/?demo=1#binder'],
    ['Projects App', base+'/projects-app/?demo=1'],
    ['Bay Map Builder', base+'/bay-map-builder/']
  ];
  const current = location.pathname;
  const css = document.createElement('style');
  css.textContent = `
    #demoHubNav,#demoHubNav *{box-sizing:border-box;font-family:Inter,"Segoe UI",Arial,sans-serif}
    #demoHubNav{position:fixed!important;inset:0 auto 0 0!important;width:224px!important;background:#163b30!important;color:#edf5ef!important;z-index:2147483000!important;overflow-y:auto!important;box-shadow:3px 0 15px #162a1c2b}
    #demoHubNav .dh-head{padding:21px 15px 15px;border-bottom:1px solid #ffffff24}
    #demoHubNav .dh-head b{font-size:16px;color:white}
    #demoHubNav .dh-head small{display:block;margin-top:5px;color:#b9d1c4;font-size:11px;line-height:1.4}
    #demoHubNav nav{display:flex;flex-direction:column;gap:3px;padding:12px 9px;box-shadow:none!important;background:transparent!important;position:static!important;width:auto!important;height:auto!important}
    #demoHubNav nav a{display:block;text-decoration:none!important;color:#dceae1!important;padding:11px 10px!important;border-radius:9px;font-size:13px!important;font-weight:500!important;white-space:normal!important;border:0!important;min-height:auto!important;background:transparent!important}
    #demoHubNav nav a:hover,#demoHubNav nav a[aria-current=page]{background:#346a56!important;color:#fff!important}
    #demoHubNav .dh-bottom{padding:13px;border-top:1px solid #ffffff24}
    #demoHubNav .dh-bottom a{color:#d9e9dc!important;text-decoration:none;font-size:12px}
    #demoHubToggle{position:fixed!important;top:12px;left:232px;z-index:2147483001!important;background:#163b30!important;color:#fff!important;border:0!important;border-radius:0 8px 8px 0!important;width:32px;height:40px;font:20px sans-serif!important;cursor:pointer!important}
    body.demo-hub-visible{margin-left:224px!important;width:calc(100% - 224px)!important;max-width:none!important}
    body.demo-hub-hidden{margin-left:0!important;width:100%!important}
    body.demo-hub-hidden #demoHubNav{transform:translateX(-100%)!important}
    body.demo-hub-hidden #demoHubToggle{left:0!important}
    @media(max-width:820px){body.demo-hub-visible{margin-left:0!important;width:100%!important}#demoHubToggle{left:224px!important}}
  `;
  document.head.appendChild(css);
  const side = document.createElement('aside');
  side.id='demoHubNav';
  side.setAttribute('aria-label','Demo library navigation');
  const safe = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  side.innerHTML='<div class="dh-head"><b>Demo Library</b><small>Original demos · Switch projects</small></div><nav aria-label="Product demos">'+items.map(([label,url])=>'<a href="'+safe(url)+'"'+(new URL(url).pathname===current?' aria-current="page"':'')+'>'+safe(label)+'</a>').join('')+'</nav><div class="dh-bottom"><a href="'+base+'/projects-app/showcase/">← All demos</a></div>';
  const toggle=document.createElement('button');
  toggle.id='demoHubToggle'; toggle.type='button';toggle.setAttribute('aria-label','Toggle demo sidebar');toggle.textContent='‹';
  toggle.addEventListener('click',()=>{
    const hide=document.body.classList.toggle('demo-hub-hidden');
    document.body.classList.toggle('demo-hub-visible',!hide);
    toggle.textContent=hide?'☰':'‹';
    toggle.setAttribute('aria-expanded',String(!hide));
  });
  const mount=()=>{
    if(document.getElementById('demoHubNav'))return;
    document.body.append(side,toggle);
    document.body.classList.add('demo-hub-visible');
    if(window.matchMedia('(max-width:820px)').matches)toggle.click();
  };
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount,{once:true});
  else mount();
})();