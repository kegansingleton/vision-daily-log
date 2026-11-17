/* ===== Vision portal: "Report a problem" button on every page =====
   Saves to the SharePoint list "Portal Feedback" (a Power Automate flow emails it).
   Uses its own MSAL instance on the same app + cache, so it works on any page the person is signed in to. */
(function(){
  if(window.__vmFeedback) return; window.__vmFeedback=1;
  var SITE="visionmech.sharepoint.com,ccfe4256-b2bf-4a3b-8e37-50f8021bb3b2,acce7714-1643-4364-86e0-8641745e0281";
  var SITEURL="https://graph.microsoft.com/v1.0/sites/"+SITE, SCOPE="https://graph.microsoft.com/Sites.ReadWrite.All", LIST="Portal Feedback";
  var H2C="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
  var errs=[]; window.addEventListener("error",function(e){ errs.push(String(e.message||"")+" @"+(e.lineno||"")); if(errs.length>5) errs.shift(); });
  function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  function el(tag,css,html){ var x=document.createElement(tag); if(css) x.style.cssText=css; if(html!=null) x.innerHTML=html; return x; }
  function ready(fn){ if(document.readyState!=="loading") fn(); else document.addEventListener("DOMContentLoaded",fn); }
  var app=null;
  /* ===== Guest / viewer mode: these logins can open everything but nothing they do is saved ===== */
  var VIEWERS=["pguest@visionmech.net"];
  var GUEST_MSG="Cannot submit — you are viewing as a guest.";
  function isViewer(){ var ac=account(); return !!(ac&&VIEWERS.indexOf(String(ac.username||"").toLowerCase())>-1); }
  var FB_SENDING=false;
  function toast(t){ var d=el("div","position:fixed;left:50%;top:18px;transform:translateX(-50%);z-index:9999;background:#1f2937;color:#fff;padding:12px 18px;border-radius:10px;font:700 14px -apple-system,Segoe UI,Roboto,Arial,sans-serif;box-shadow:0 6px 20px rgba(0,0,0,.3);max-width:90%;text-align:center",esc(t)); d.setAttribute("data-html2canvas-ignore","1"); document.body.appendChild(d); setTimeout(function(){ d.remove(); },3500); }
  var _fetch=window.fetch.bind(window);
  window.fetch=function(input,init){
    try{ var url=typeof input==="string"?input:(input&&input.url)||"", m=String((init&&init.method)||(input&&input.method)||"GET").toUpperCase();
      if(m!=="GET"&&m!=="HEAD"&&/graph\.microsoft\.com|sharepoint\.com/i.test(url)&&!FB_SENDING&&isViewer()){ toast(GUEST_MSG); return Promise.reject(new Error(GUEST_MSG)); } }catch(e){}
    return _fetch(input,init); };
  ready(function(){ if(!isViewer()) return; var b=el("div","position:fixed;right:10px;bottom:88px;z-index:9990;background:#1d4ed8;color:#fff;border-radius:18px;padding:6px 12px;font:700 12px -apple-system,Segoe UI,Roboto,Arial,sans-serif;opacity:.9","&#128065; Viewing as guest — read only"); b.setAttribute("data-html2canvas-ignore","1"); document.body.appendChild(b); });
  function msalApp(){ if(app) return app; if(!window.msal) return null;
    app=new msal.PublicClientApplication({auth:{clientId:"ece3c8e3-0f54-49dc-bf83-b46411ade0aa",authority:"https://login.microsoftonline.com/8d64d0b9-91ba-4959-a83c-2e7e8cf233b3",redirectUri:(location.origin+location.pathname)},cache:{cacheLocation:"localStorage"}}); return app; }
  function account(){ var a=msalApp(); if(!a) return null; var all=a.getAllAccounts(); return a.getActiveAccount()||all.filter(function(x){return /@visionmech\.net$/i.test(x.username||"");})[0]||all[0]||null; }
  async function token(){ var a=msalApp(), ac=account(); if(!a||!ac) throw new Error("not signed in"); return (await a.acquireTokenSilent({scopes:[SCOPE],account:ac})).accessToken; }
  function loadH2C(){ return new Promise(function(res){ if(window.html2canvas){ res(window.html2canvas); return; } var s=document.createElement("script"); s.src=H2C; s.onload=function(){ res(window.html2canvas||null); }; s.onerror=function(){ res(null); }; document.head.appendChild(s); }); }
  /* ===== Installable app: register the offline helper + show an "Install" banner on the portal ===== */
  if("serviceWorker" in navigator && location.protocol==="https:"){ window.addEventListener("load",function(){ navigator.serviceWorker.register("/sw.js").catch(function(){}); }); }
  var deferredInstall=null;
  window.addEventListener("beforeinstallprompt",function(e){ e.preventDefault(); deferredInstall=e; ready(installBanner); });
  function standalone(){ return (window.matchMedia&&window.matchMedia("(display-mode: standalone)").matches)||window.navigator.standalone===true; }
  function isIOS(){ return /iphone|ipad|ipod/i.test(navigator.userAgent)||(navigator.platform==="MacIntel"&&navigator.maxTouchPoints>1); }
  function dismissed(){ try{ return localStorage.getItem("vm_install_dismissed")==="1"; }catch(e){ return false; } }
  function installBanner(){
    if(standalone()||dismissed()||document.getElementById("vmInstall")) return;
    if(!/\/(portal\.html)?$/.test(location.pathname)) return;               /* only on the portal home page */
    var ios=isIOS()&&!deferredInstall; if(!deferredInstall&&!ios) return;
    var b=el("div","position:fixed;left:10px;right:10px;bottom:14px;z-index:9991;background:#fff;color:#1f2937;border-radius:14px;box-shadow:0 8px 28px rgba(0,0,0,.28);padding:12px 14px;display:flex;gap:12px;align-items:center;font:14px -apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:520px;margin:0 auto");
    b.id="vmInstall"; b.setAttribute("data-html2canvas-ignore","1");
    b.innerHTML='<img src="/icon-192.png" width="44" height="44" style="border-radius:10px;flex:0 0 44px" alt=""><div style="flex:1"><b>Put Vision on your home screen</b><div style="font-size:12.5px;color:#6b7280;margin-top:2px">'+
      (ios?'Tap the <b>Share</b> button <span style="font-size:15px">&#x2B06;&#xFE0E;</span> at the bottom of Safari, then <b>Add to Home Screen</b>.':'Opens like an app — full screen, faster on weak signal.')+'</div></div>'+
      (ios?'':'<button id="vmInstGo" style="background:#c0392b;color:#fff;border:none;border-radius:9px;padding:10px 14px;font-weight:800;font-size:14px">Install</button>')+
      '<button id="vmInstX" aria-label="Close" style="background:none;border:none;font-size:22px;color:#9ca3af;padding:0 2px">&times;</button>';
    document.body.appendChild(b);
    document.getElementById("vmInstX").onclick=function(){ try{ localStorage.setItem("vm_install_dismissed","1"); }catch(e){} b.remove(); };
    var go=document.getElementById("vmInstGo"); if(go) go.onclick=function(){ if(!deferredInstall) return; deferredInstall.prompt(); deferredInstall.userChoice.then(function(){ deferredInstall=null; b.remove(); }); };
  }
  ready(function(){ if(isIOS()&&!standalone()&&/safari/i.test(navigator.userAgent)&&!/crios|fxios|edgios/i.test(navigator.userAgent)) installBanner(); });
  var shot="";
  ready(function(){
    if(/[?&]print=/.test(location.search)) return;
    var btn=el("button","position:fixed;left:10px;bottom:88px;z-index:9990;background:rgba(127,29,29,.9);color:#fff;border:none;border-radius:18px;padding:7px 12px;font:700 12px -apple-system,Segoe UI,Roboto,Arial,sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.25);cursor:pointer;opacity:.85","&#9873; Problem?");
    btn.title="Report a problem with this page"; btn.setAttribute("data-html2canvas-ignore","1"); document.body.appendChild(btn);
    var ov=el("div","position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9995;display:none;padding:14px;overflow:auto;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif");
    ov.setAttribute("data-html2canvas-ignore","1");
    ov.innerHTML='<div style="background:#fff;max-width:480px;margin:30px auto;border-radius:12px;padding:16px;color:#1f2937">'+
      '<div style="font-size:17px;font-weight:800;color:#7f1d1d;margin-bottom:4px">Report a problem</div>'+
      '<div style="font-size:12.5px;color:#6b7280;margin-bottom:8px">Tell us what went wrong. The page, your name, your device and a screenshot are sent automatically.</div>'+
      '<select id="vmfbCat" style="width:100%;padding:10px;font-size:15px;border:1.5px solid #9ca3af;border-radius:9px;margin-bottom:8px"><option>Something is broken / error</option><option>Can\'t find or pick something</option><option>Wrong info showing</option><option>Idea / request</option><option>Other</option></select>'+
      '<textarea id="vmfbTxt" rows="4" placeholder="What were you trying to do, and what happened?" style="width:100%;padding:10px;font-size:15px;border:1.5px solid #9ca3af;border-radius:9px;font-family:inherit;box-sizing:border-box"></textarea>'+
      '<label style="display:flex;gap:8px;align-items:center;font-size:13.5px;margin:8px 0;color:#1f2937;font-weight:600"><input type="checkbox" id="vmfbShot" checked style="width:18px;height:18px;flex:0 0 18px;margin:0;padding:0"> <span>Include a screenshot of the page</span></label>'+
      '<div id="vmfbPrev" style="display:none;margin:6px 0"><img id="vmfbImg" style="max-width:100%;max-height:160px;border:1px solid #e5e7eb;border-radius:8px"></div>'+
      '<div id="vmfbMsg" style="font-size:13.5px;margin:8px 0;display:none"></div>'+
      '<div style="display:flex;gap:8px;margin-top:10px"><button id="vmfbCancel" style="flex:1;padding:11px;border-radius:9px;border:1.5px solid #c0392b;background:#fff;color:#c0392b;font-weight:800;cursor:pointer">Cancel</button><button id="vmfbSend" style="flex:1;padding:11px;border-radius:9px;border:none;background:#c0392b;color:#fff;font-weight:800;cursor:pointer">Send</button></div></div>';
    document.body.appendChild(ov);
    function $(id){ return document.getElementById(id); }
    function msg(t,ok){ var m=$("vmfbMsg"); m.style.display="block"; m.style.color=ok?"#15803d":"#b91c1c"; m.textContent=t; }
    btn.onclick=async function(){ shot=""; $("vmfbMsg").style.display="none"; $("vmfbTxt").value=""; $("vmfbSend").disabled=false; $("vmfbSend").textContent="Send"; $("vmfbPrev").style.display="none";
      btn.style.display="none";
      try{ var h=await loadH2C(); if(h){ var c=await h(document.body,{scale:Math.min(1,900/Math.max(1,window.innerWidth)),useCORS:true,logging:false,windowHeight:window.innerHeight,height:window.innerHeight,y:window.scrollY}); shot=c.toDataURL("image/jpeg",0.6); $("vmfbImg").src=shot; $("vmfbPrev").style.display="block"; } }catch(e){ shot=""; }
      btn.style.display=""; ov.style.display="block"; $("vmfbTxt").focus(); };
    $("vmfbCancel").onclick=function(){ ov.style.display="none"; };
    $("vmfbSend").onclick=async function(){ var txt=$("vmfbTxt").value.trim(); if(!txt){ msg("Type a quick description first."); return; }
      var b=$("vmfbSend"); b.disabled=true; b.textContent="Sending…"; FB_SENDING=true;
      var ac=account(), page=(document.title||"").replace(/^Vision Mechanical\s*[—-]\s*/,"")||location.pathname;
      var f={Title:page+" | "+(ac?(ac.name||ac.username):"not signed in"), Page:page, PageUrl:location.href.split("#")[0], Category:$("vmfbCat").value, Details:txt,
        ReportedBy:ac?(ac.name||""):"", LoginEmail:ac?String(ac.username||"").toLowerCase():"", Device:navigator.userAgent+" · "+window.innerWidth+"x"+window.innerHeight+(errs.length?" · recent errors: "+errs.join(" | "):""),
        Screenshot:$("vmfbShot").checked?shot:"", Status:"New", SubmittedAt:new Date().toLocaleString()};
      try{ var t=await token(), L=await fetch(SITEURL+"/lists?$select=id,displayName&$top=999",{headers:{Authorization:"Bearer "+t}}).then(function(r){return r.json();});
        var l=(L.value||[]).filter(function(x){return (x.displayName||"").toLowerCase()===LIST.toLowerCase();})[0]; if(!l) throw new Error("feedback list missing");
        var c=await fetch(SITEURL+"/lists/"+l.id+"/columns?$select=name,readOnly",{headers:{Authorization:"Bearer "+t}}).then(function(r){return r.json();}), cols={}; (c.value||[]).forEach(function(x){ if(!x.readOnly) cols[x.name]=1; });
        Object.keys(f).forEach(function(k){ if(k!=="Title"&&!cols[k]) delete f[k]; if(f[k]==="") delete f[k]; });
        var r=await fetch(SITEURL+"/lists/"+l.id+"/items",{method:"POST",headers:{Authorization:"Bearer "+t,"Content-Type":"application/json"},body:JSON.stringify({fields:f})});
        if(!r.ok) throw new Error("HTTP "+r.status);
        FB_SENDING=false; msg("Sent — thanks. The office will take a look.",true); b.textContent="Sent ✓"; setTimeout(function(){ ov.style.display="none"; },1600);
      }catch(e){ FB_SENDING=false; b.disabled=false; b.textContent="Send";
        var body="Page: "+location.href+"\n"+$("vmfbCat").value+"\n\n"+txt+"\n\n"+navigator.userAgent;
        msg("Couldn't send it from here ("+e.message+"). Tap here to email it instead.",false);
        $("vmfbMsg").style.cursor="pointer"; $("vmfbMsg").onclick=function(){ location.href="mailto:reports@visionmech.net?subject="+encodeURIComponent("Portal problem: "+page)+"&body="+encodeURIComponent(body); }; } };
  });
})();
