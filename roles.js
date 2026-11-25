/* ===== Vision roles: who is an Owner, a PM or a Service tech =====
   One place for it: the Owner / PM / Service checkboxes in Employee Setup (Employee Master list).
   Every page reads window.VMR right away (from this device's cache, or the built-in defaults the first time),
   then this script quietly refreshes the cache from SharePoint in the background for the next page.
   Foreman, Admin and Shop stay where they always were (checkboxes the pages read directly). */
(function(){
  var KEY="vm_rolelist_v1";
  var DEF={owners:["kegan.singleton@visionmech.net","bsingleton@visionmech.net"],
           pms:["avrell@visionmech.net","jbeach@visionmech.net","sthomas@visionmech.net","avigil@visionmech.net"],
           svc:["kharvey@visionmech.net","jbarron@visionmech.net"]};
  var c=null; try{ c=JSON.parse(localStorage.getItem(KEY)||"null"); }catch(e){}
  function pick(k){ return (c&&Array.isArray(c[k])&&c[k].length)?c[k].slice():DEF[k].slice(); }
  var R={owners:pick("owners"), pms:pick("pms"), svc:(c&&Array.isArray(c.svc))?c.svc.slice():DEF.svc.slice()};
  R.isOwner=function(e){ return R.owners.indexOf(String(e||"").toLowerCase())>-1; };
  R.isPM=function(e){ return R.pms.indexOf(String(e||"").toLowerCase())>-1; };
  R.isSvc=function(e){ return R.svc.indexOf(String(e||"").toLowerCase())>-1; };
  /* pages that already read the Employee Master can hand their rows in to keep the cache fresh */
  R.store=function(rows){ try{
    var y=function(v){ return v===true||v==="Yes"||v==="true"; }, out={owners:[],pms:[],svc:[],t:Date.now()};
    (rows||[]).forEach(function(f){ if(!f||(f.Active||"Yes")==="No") return; var e=String(f.LoginEmail||"").toLowerCase().trim()||auto(f.EmployeeName); if(!e) return;
      if(y(f.Owner)) out.owners.push(e); if(y(f.PM)) out.pms.push(e); if(y(f.ServiceTech)) out.svc.push(e); });
    if(!out.owners.length) return;            /* never cache an empty owner list (a half-loaded read) */
    localStorage.setItem(KEY,JSON.stringify(out)); }catch(e){} };
  function auto(n){ var p=String(n||"").trim().split(/\s+/).filter(Boolean); if(!p.length) return ""; var fi=(p[0][0]||"").toLowerCase(), ln=p[p.length-1].toLowerCase().replace(/[^a-z0-9]/g,""); return (fi&&ln)?(fi+ln+"@visionmech.net"):""; }
  window.VMR=R;

  async function refresh(){
    if(!window.msal||!window.fetch) return;
    if(c&&c.t&&Date.now()-c.t<15*60*1000) return;   /* fresh enough */
    try{
      var app=new msal.PublicClientApplication({auth:{clientId:"ece3c8e3-0f54-49dc-bf83-b46411ade0aa",authority:"https://login.microsoftonline.com/8d64d0b9-91ba-4959-a83c-2e7e8cf233b3",redirectUri:location.origin+location.pathname},cache:{cacheLocation:"localStorage"}});
      var all=app.getAllAccounts()||[], acct=all.filter(function(a){return /@visionmech\.net$/i.test(a.username||"");})[0]||all[0]; if(!acct) return;
      var tok=(await app.acquireTokenSilent({scopes:["https://graph.microsoft.com/Sites.ReadWrite.All"],account:acct})).accessToken;
      var url="https://graph.microsoft.com/v1.0/sites/visionmech.sharepoint.com,ccfe4256-b2bf-4a3b-8e37-50f8021bb3b2,acce7714-1643-4364-86e0-8641745e0281/lists/537b8d9b-e666-43d1-9441-efc2bed48d73/items?$expand=fields($select=EmployeeName,LoginEmail,Owner,PM,ServiceTech,Active)&$top=999", rows=[];
      while(url){ var j=await fetch(url,{headers:{Authorization:"Bearer "+tok}}).then(function(r){return r.json();}); if(j.error) return; (j.value||[]).forEach(function(it){ if(it.fields) rows.push(it.fields); }); url=j["@odata.nextLink"]||null; }
      R.store(rows);
    }catch(e){}
  }
  window.addEventListener("load",function(){ setTimeout(refresh,2500); });
})();
