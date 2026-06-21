"use strict";
(function(){
  /* ===== configuración ===== */
  /* API SODA de Socrata (datos.gov.co): soporta CORS desde el navegador.
     El endpoint OData v4 (/api/odata/v4/...) no envía cabeceras CORS y
     provoca "Failed to fetch" al consultarlo desde una página web. */
  var API = "https://www.datos.gov.co/resource/jbjy-vk9h.json";
  var PAGE = 20;

  /* documentos de proveedor restringidos: no se pueden buscar y se excluyen
     de todos los resultados (filtro != en el lado del servidor, SoQL válido). */
  var BLOCKED_DOCS = ["1128272022"];

  /* candidatos por campo lógico (se resuelven contra una fila de muestra) */
  var FIELD_CANDIDATES = {
    entidad:      ["nombre_entidad","nombre_de_la_entidad","nombreentidad","entidad"],
    nitEntidad:   ["nit_entidad","nit_de_la_entidad","nit_entidad_contratante"],
    proveedor:    ["proveedor_adjudicado","nombre_del_proveedor","nombre_proveedor","proveedor"],
    docProveedor: ["documento_proveedor","nit_proveedor","documento_del_proveedor","documento"],
    objeto:       ["objeto_del_contrato","objeto_a_contratar","objeto","descripcion_del_proceso"],
    descripcion:  ["descripcion_del_proceso","descripci_n_del_proceso","descripcion"],
    referencia:   ["referencia_del_contrato","referencia_del_proceso","referencia","id_contrato"],
    repLegalId:   ["identificaci_n_representante_legal","identificacion_representante_legal","nit_representante_legal","id_representante_legal"],
    repLegal:     ["nombre_representante_legal","representante_legal"],
    modalidad:    ["modalidad_de_contratacion","tipo_de_proceso","modalidad"],
    valor:        ["valor_del_contrato","valor_total_del_contrato","valor_contrato","valor"],
    fechaFirma:   ["fecha_de_firma","fecha_de_firma_del_contrato","fecha_firma"],
    url:          ["urlproceso","url_proceso","url"],
    departamento: ["departamento","departamento_entidad","dpto"],
    ciudad:       ["ciudad","ciudad_entidad","municipio"],
    estado:       ["estado_contrato","estado_del_contrato","estado"]
  };

  var MODALIDADES = ["Contratación directa","Mínima cuantía","Licitación pública",
    "Selección abreviada de menor cuantía","Selección abreviada subasta inversa",
    "Concurso de méritos abierto","Concurso de méritos con lista corta",
    "Contratación régimen especial","Contratación régimen especial (con ofertas)",
    "Asociación Público Privada","Enajenación de bienes con sobre cerrado","Enajenación de bienes con subasta"];

  function modTheme(m){
    var s=(m||"").toLowerCase();
    if(s.indexOf("directa")>=0) return ["#0E6E78","#E2F0F1","#0A555D"];
    if(s.indexOf("mínima")>=0||s.indexOf("minima")>=0) return ["#B98A28","#F3E8CE","#7A5C16"];
    if(s.indexOf("licitaci")>=0) return ["#3457A6","#E4EAF7","#243f78"];
    if(s.indexOf("abreviada")>=0) return ["#7A4FB0","#EEE6F6","#553584"];
    if(s.indexOf("méritos")>=0||s.indexOf("meritos")>=0) return ["#1E7A5A","#DDF0E7","#125640"];
    if(s.indexOf("especial")>=0) return ["#B4452F","#F6E1DB","#822f1f"];
    return ["#16314f","#E7ECF3","#0A1A2F"];
  }

  var COP=new Intl.NumberFormat("es-CO",{style:"currency",currency:"COP",maximumFractionDigits:0});
  var NUM=new Intl.NumberFormat("es-CO");

  function esc(s){ return String(s).replace(/'/g,"''"); }
  function escHtml(s){ return String(s==null?"":s)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"); }
  function pick(keys,cands){ var low={}; keys.forEach(function(k){ low[k.toLowerCase()]=k; });
    for(var i=0;i<cands.length;i++){ var c=cands[i].toLowerCase(); if(low[c]) return low[c]; } return null; }
  function val(row,key){ if(!key) return undefined; var v=row[key];
    if(v&&typeof v==="object") return v.url||v.description||v.label||undefined; return v; }
  function fmtFecha(v){ if(!v) return null; var d=new Date(v); if(isNaN(d)) return String(v).slice(0,10);
    return d.toLocaleDateString("es-CO",{year:"numeric",month:"short",day:"2-digit"}); }
  function getYear(v){ if(!v) return null; var d=new Date(v); return isNaN(d)?null:d.getFullYear(); }

  /* ===== estado ===== */
  var F=null, active=null, rows=[], count=null, page=0,
      loading=false, more=false, done=false, error=null, reqId=0, xlsBusy=false;
  var $=function(id){ return document.getElementById(id); };
  var list=$("list"), rcount=$("rcount"), rsub=$("rsub");

  /* poblar desplegables finitos */
  (function(){
    var m=$("f_mod"); MODALIDADES.forEach(function(x){ var o=document.createElement("option"); o.value=x;o.textContent=x;m.appendChild(o); });
    var y=$("f_anio"), now=new Date().getFullYear();
    for(var yr=now; yr>=2015; yr--){ var o=document.createElement("option"); o.value=yr;o.textContent=yr;y.appendChild(o); }
  })();

  /* ===== consultas SODA (SoQL) ===== */
  /* opts.lower    -> usa lower()+like en campos de texto (case-insensitive)
     opts.idNumeric-> trata los identificadores (NIT/documento) como número.
     Los campos NIT/documento suelen ser columnas numéricas: aplicarles
     lower()/like provoca HTTP 400, por eso usan coincidencia exacta. */
  function buildWhere(a, opts){
    if(!F) return "";
    var p=[];
    function txt(field, value){          /* campos de texto libre: contiene */
      if(!value || !field) return;
      var v=esc(opts.lower? value.toLowerCase() : value);
      p.push(opts.lower? "lower("+field+") like '%"+v+"%'" : field+" like '%"+v+"%'");
    }
    function id(field, value){            /* identificadores: coincidencia exacta */
      if(!value || !field) return;
      var raw=String(value).trim();
      if(opts.idNumeric && /^[0-9]+$/.test(raw)) p.push(field+" = "+raw);   /* columna numérica */
      else p.push(field+" = '"+esc(raw)+"'");                               /* columna de texto */
    }
    id(F.nitEntidad,    a.nitEnt);
    txt(F.entidad,      a.nomEnt);
    txt(F.objeto,       a.objeto);
    txt(F.referencia,   a.ref);
    txt(F.descripcion,  a.desc);
    id(F.docProveedor,  a.nitProv);
    txt(F.proveedor,    a.nomProv);
    id(F.repLegalId,    a.nitRep);
    if(a.mod && F.modalidad) p.push(F.modalidad+" = '"+esc(a.mod)+"'");
    if(a.anio && F.fechaFirma){ var yy=parseInt(a.anio,10);
      p.push(F.fechaFirma+" >= '"+yy+"-01-01T00:00:00' and "+F.fechaFirma+" < '"+(yy+1)+"-01-01T00:00:00'"); }
    if(F.docProveedor){                  /* excluir proveedores restringidos */
      BLOCKED_DOCS.forEach(function(doc){
        if(opts.idNumeric && /^[0-9]+$/.test(doc)) p.push(F.docProveedor+" != "+doc);
        else p.push(F.docProveedor+" != '"+esc(doc)+"'");
      });
    }
    return p.join(" and ");
  }
  /* orden: siempre por fecha de firma (más reciente primero) y los contratos
     SIN fecha de firma al final. coalesce empuja los vacíos a una fecha muy
     antigua; si el servidor no soporta coalesce, se usa el orden simple. */
  function orderCandidates(){
    if(!(F && F.fechaFirma)) return [""];
    var f=F.fechaFirma;
    return ["coalesce("+f+",'1111-01-01T00:00:00') desc", f+" desc"];
  }
  /* estrategias de respaldo (orden × tipo de filtro) que se prueban ante 400 */
  var WHERE_OPTS=[{lower:true,idNumeric:true},{lower:true,idNumeric:false},{lower:false,idNumeric:false}];
  function buildStrategies(){
    var s=[]; orderCandidates().forEach(function(ord){
      WHERE_OPTS.forEach(function(opts){ s.push({ord:ord, opts:opts}); });
    }); return s;
  }
  function tryStrategies(attempt){
    return buildStrategies().reduce(function(promise, st, i){
      return i===0 ? attempt(st) : promise.catch(function(){ return attempt(st); });
    }, null);
  }
  function fetchPage(a, idx, withCount){
    var off=idx*PAGE;
    function attempt(st){
      var w=buildWhere(a,st.opts);
      var dataUrl=API+"?$limit="+PAGE+"&$offset="+off;
      if(w) dataUrl+="&$where="+encodeURIComponent(w);
      if(st.ord) dataUrl+="&$order="+encodeURIComponent(st.ord);
      var pData=fetch(dataUrl).then(function(r){ if(!r.ok) throw new Error("HTTP "+r.status); return r.json(); });
      if(!withCount) return pData.then(function(arr){ return {value:arr}; });
      var cntUrl=API+"?$select="+encodeURIComponent("count(1) as cnt")+(w? "&$where="+encodeURIComponent(w):"");
      var pCount=fetch(cntUrl).then(function(r){ return r.ok?r.json():null; }).catch(function(){ return null; });
      return Promise.all([pData,pCount]).then(function(res){
        var out={value:res[0]};
        if(res[1] && res[1][0] && res[1][0].cnt!=null) out["@odata.count"]=Number(res[1][0].cnt);
        return out;
      });
    }
    return tryStrategies(attempt);
  }

  /* ===== descarga a Excel ===== */
  var XLS_CAP=5000, XLS_BATCH=1000;   /* tope de filas y tamaño de lote */
  var XLS_COLS=[
    ["entidad","Entidad"], ["nitEntidad","NIT entidad"],
    ["proveedor","Proveedor"], ["docProveedor","Documento proveedor"],
    ["objeto","Objeto del contrato"], ["referencia","Referencia"],
    ["modalidad","Modalidad"], ["valor","Valor del contrato"],
    ["fechaFirma","Fecha de firma"], ["departamento","Departamento"],
    ["ciudad","Ciudad"], ["estado","Estado"],
    ["repLegal","Representante legal"], ["repLegalId","Doc. rep. legal"],
    ["url","URL del proceso"]
  ];
  /* trae solo datos (sin conteo) en lotes grandes, con la misma lógica de respaldo */
  function fetchRows(a, offset, limit){
    function attempt(st){
      var w=buildWhere(a,st.opts);
      var url=API+"?$limit="+limit+"&$offset="+offset;
      if(w) url+="&$where="+encodeURIComponent(w);
      if(st.ord) url+="&$order="+encodeURIComponent(st.ord);
      return fetch(url).then(function(r){ if(!r.ok) throw new Error("HTTP "+r.status); return r.json(); });
    }
    return tryStrategies(attempt);
  }
  function buildRecords(allRows){
    return allRows.map(function(row){
      var rec={};
      XLS_COLS.forEach(function(c){
        var v=val(row, F[c[0]]);
        if(c[0]==="valor"){ var n=(v!=null&&v!=="")?Number(v):null; rec[c[1]]=(n!=null&&!isNaN(n))?n:(v||""); }
        else if(c[0]==="fechaFirma"){ rec[c[1]]=v?String(v).slice(0,10):""; }
        else rec[c[1]]=(v==null)?"":v;
      });
      return rec;
    });
  }
  function triggerDownload(blob, fname){
    var url=URL.createObjectURL(blob), el=document.createElement("a");
    el.href=url; el.download=fname; document.body.appendChild(el); el.click();
    document.body.removeChild(el); setTimeout(function(){ URL.revokeObjectURL(url); }, 1500);
  }
  function exportRecords(records, base){
    if(window.XLSX){
      var ws=XLSX.utils.json_to_sheet(records);
      var wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Contratos");
      XLSX.writeFile(wb, base+".xlsx");
    } else {
      var headers=Object.keys(records[0]);
      function cell(v){ return '"'+String(v==null?"":v).replace(/"/g,'""')+'"'; }
      var lines=[headers.map(cell).join(",")];
      records.forEach(function(r){ lines.push(headers.map(function(h){ return cell(r[h]); }).join(",")); });
      triggerDownload(new Blob(["﻿"+lines.join("\r\n")], {type:"text/csv;charset=utf-8;"}), base+".csv");
    }
  }
  function setXlsBusy(on){
    var b=$("btnXlsx"); if(!b) return;
    b.disabled = on || !F;
    b.lastChild.textContent = on? " Generando…" : " Descargar Excel";
  }
  function downloadExcel(){
    if(!F || xlsBusy) return;
    var a=active||readForm();
    if(isBlocked(a)){ showBlocked(); return; }
    xlsBusy=true; setXlsBusy(true);
    var all=[];
    function loop(offset){
      var lim=Math.min(XLS_BATCH, XLS_CAP-all.length);
      if(lim<=0) return Promise.resolve();
      return fetchRows(a, offset, lim).then(function(arr){
        arr=arr||[]; all=all.concat(arr);
        if(arr.length<lim || all.length>=XLS_CAP) return;
        return loop(offset+arr.length);
      });
    }
    loop(0).then(function(){
      if(!all.length){ alert("No hay contratos para descargar con los filtros actuales."); return; }
      exportRecords(buildRecords(all), "contratos-secop-ii-"+new Date().toISOString().slice(0,10));
    }).catch(function(e){
      alert("No se pudo generar el archivo: "+((e&&e.message)||"error de red"));
    }).then(function(){ xlsBusy=false; setXlsBusy(false); });
  }

  /* ===== render ===== */
  function setBtnLoading(on){
    $("btnIcon").innerHTML = on
      ? '<span class="spin"></span>'
      : '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>';
    $("btnSearch").disabled = on || !F;
  }

  function cardHtml(row){
    var ent=val(row,F.entidad), nitE=val(row,F.nitEntidad), prov=val(row,F.proveedor),
        docP=val(row,F.docProveedor), obj=val(row,F.objeto)||val(row,F.descripcion),
        ref=val(row,F.referencia), mod=val(row,F.modalidad), valor=val(row,F.valor),
        fecha=val(row,F.fechaFirma), url=val(row,F.url), depto=val(row,F.departamento),
        ciudad=val(row,F.ciudad), estado=val(row,F.estado),
        rep=val(row,F.repLegal), repId=val(row,F.repLegalId);

    var th=modTheme(mod), accent=th[0], soft=th[1], aink=th[2];
    var year=getYear(fecha);
    var valNum=(valor!=null && valor!=="")? Number(valor):null;
    var loc=[ciudad,depto].filter(Boolean).join(", ");

    var meta="";
    if(nitE) meta+='<span><span class="k">NIT</span> '+escHtml(nitE)+'</span>';
    if(loc)  meta+='<span>'+escHtml(loc)+'</span>';
    if(ref)  meta+='<span><span class="k">Ref.</span> '+escHtml(ref)+'</span>';

    var chips="";
    if(mod)    chips+='<span class="chip mod">'+escHtml(mod)+'</span>';
    if(estado) chips+='<span class="chip">'+escHtml(estado)+'</span>';
    if(fecha)  chips+='<span class="chip">Firmado '+escHtml(fmtFecha(fecha))+'</span>';

    var provHtml="";
    if(prov){ provHtml='<div class="prov"><span class="plabel">Proveedor</span>'+
      '<span class="who">'+escHtml(prov)+'</span>'+(docP? '<span class="nit">· '+escHtml(docP)+'</span>':'')+'</div>'; }

    var repHtml="";
    if(rep||repId){ repHtml='<div class="replegal">Rep. legal: '+escHtml(rep||"—")+
      (repId? ' <span class="nit">('+escHtml(repId)+')</span>':'')+'</div>'; }

    var safeUrl=(url && /^https?:\/\//i.test(url))? url : null;
    var verproc = safeUrl
      ? '<a class="verproc" href="'+escHtml(safeUrl)+'" target="_blank" rel="noopener noreferrer">Ver proceso '+
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M7 17 17 7M9 7h8v8"/></svg></a>'
      : '<span class="verproc disabled">Sin enlace</span>';

    return ''+
    '<article class="card" style="--accent:'+accent+';--accent-soft:'+soft+';--accent-ink:'+aink+'">'+
      '<div class="card-grid">'+
        '<div class="card-main">'+
          '<h2 class="ent">'+escHtml(ent||"Entidad no registrada")+'</h2>'+
          (meta? '<div class="meta-line">'+meta+'</div>':'')+
          (obj? '<p class="objeto">'+escHtml(obj)+'</p>':'')+
          provHtml+ repHtml+
          (chips? '<div class="chips">'+chips+'</div>':'')+
        '</div>'+
        '<div class="card-side">'+
          '<div>'+
            '<div class="valor-lbl">Valor del contrato</div>'+
            '<div class="valor">'+((valNum!=null && !isNaN(valNum))? COP.format(valNum):'—')+'</div>'+
            (year? '<div class="anio">Año de firma · <b>'+year+'</b></div>':'')+
          '</div>'+ verproc+
        '</div>'+
      '</div>'+
    '</article>';
  }

  function render(){
    if(loading){ rcount.innerHTML="Buscando…"; }
    else if(count!=null){ rcount.innerHTML='<span>'+NUM.format(count)+'</span> contrato'+(count===1?'':'s')+' encontrado'+(count===1?'':'s'); }
    else { rcount.textContent=rows.length+" resultado"+(rows.length===1?'':'s'); }
    var hasF = active && (active.nitEnt||active.nomEnt||active.objeto||active.ref||active.desc||active.nitProv||active.nomProv||active.nitRep||active.mod||active.anio);
    rsub.textContent = hasF? "Según los filtros aplicados" : "Mostrando una muestra del registro nacional";

    var html="";
    if(loading && rows.length===0){ for(var i=0;i<6;i++) html+='<div class="skel"></div>'; list.innerHTML=html; return; }
    if(error){
      list.innerHTML='<div class="state error"><div class="ico">'+
        '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9 2.4 18a2 2 0 0 0 1.7 3h15.8a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/></svg>'+
        '</div><h3>La consulta no se pudo completar</h3>'+
        '<p>El servidor respondió con un error ('+escHtml(error)+'). Prueba con menos filtros o repite en unos segundos.</p></div>';
      return;
    }
    if(rows.length===0){
      list.innerHTML='<div class="state"><div class="ico">'+
        '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M4 7h16M4 12h10M4 17h7"/></svg>'+
        '</div><h3>Sin coincidencias</h3>'+
        '<p>Ningún contrato coincide con esos filtros. Prueba términos más generales o limpia alguno.</p></div>';
      return;
    }
    for(var j=0;j<rows.length;j++) html+=cardHtml(rows[j]);
    if(!done){ html+='<div class="loadmore"><button class="btn" id="loadmore"'+(more?' disabled':'')+'>'+
      (more? '<span class="spin dark"></span>Cargando…' : 'Cargar más resultados')+'</button></div>'; }
    list.innerHTML=html;
    var lm=$("loadmore"); if(lm) lm.addEventListener("click", loadMore);
  }

  /* ===== acciones ===== */
  function runQuery(a){
    active=a; var id=++reqId;
    loading=true; error=null; rows=[]; count=null; page=0; done=false;
    setBtnLoading(true); render();
    fetchPage(a,0,true).then(function(j){
      if(id!==reqId) return;
      rows=j.value||[];
      if(j["@odata.count"]!=null) count=j["@odata.count"];
      done=rows.length<PAGE;
    }).catch(function(e){ if(id===reqId) error=(e&&e.message)||"Error de consulta"; })
    .then(function(){ if(id===reqId){ loading=false; setBtnLoading(false); render(); } });
  }
  function loadMore(){
    if(more||done) return;
    var id=reqId, next=page+1; more=true; render();
    fetchPage(active,next,false).then(function(j){
      if(id!==reqId) return;
      var v=j.value||[]; rows=rows.concat(v); page=next; done=v.length<PAGE;
    }).catch(function(e){ if(id===reqId) error=(e&&e.message)||"Error"; })
    .then(function(){ if(id===reqId){ more=false; render(); } });
  }
  function readForm(){
    return {
      nitEnt:$("f_nitEnt").value.trim(), nomEnt:$("f_nomEnt").value.trim(),
      objeto:$("f_objeto").value.trim(), ref:$("f_ref").value.trim(), desc:$("f_desc").value.trim(),
      nitProv:$("f_nitProv").value.trim(), nomProv:$("f_nomProv").value.trim(),
      nitRep:$("f_nitRep").value.trim(), mod:$("f_mod").value, anio:$("f_anio").value
    };
  }
  var INPUT_IDS=["f_nitEnt","f_nomEnt","f_objeto","f_ref","f_desc","f_nitProv","f_nomProv","f_nitRep","f_mod","f_anio"];

  function showBlocked(){
    active=null; rows=[]; count=null; error=null; loading=false; done=true;
    setBtnLoading(false);
    rcount.textContent="Búsqueda no permitida";
    rsub.textContent="Este proveedor está restringido";
    list.innerHTML='<div class="state"><div class="ico">'+
      '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>'+
      '</div><h3>Proveedor restringido</h3>'+
      '<p>No es posible consultar contratos por este documento de proveedor.</p></div>';
  }
  function isBlocked(a){ return a.nitProv && BLOCKED_DOCS.indexOf(a.nitProv)>=0; }

  $("form").addEventListener("submit", function(e){ e.preventDefault(); if(!F) return;
    var a=readForm(); if(isBlocked(a)){ showBlocked(); return; } runQuery(a); });
  $("clear").addEventListener("click", function(){
    INPUT_IDS.forEach(function(id){ $(id).value=""; });
    if(F) runQuery(readForm());
  });
  $("btnXlsx").addEventListener("click", downloadExcel);

  /* campos de NIT: solo dígitos (escritura y pegado) */
  ["f_nitEnt","f_nitProv","f_nitRep"].forEach(function(id){
    $(id).addEventListener("input", function(){
      var clean=this.value.replace(/\D+/g,"");
      if(this.value!==clean) this.value=clean;
    });
  });

  /* ===== arranque ===== */
  (function boot(){
    fetch(API+"?$limit=1").then(function(r){ if(!r.ok) throw new Error("HTTP "+r.status); return r.json(); })
    .then(function(j){
      var sample=(j&&j[0])||{}; var keys=Object.keys(sample); F={};
      for(var k in FIELD_CANDIDATES) F[k]=pick(keys,FIELD_CANDIDATES[k]);
      setBtnLoading(false); setXlsBusy(false);
      fetch(API+"?$select="+encodeURIComponent("count(1) as cnt")).then(function(x){ return x.ok?x.json():null; })
        .then(function(d){ if(d && d[0] && d[0].cnt!=null) $("total").textContent=NUM.format(Number(d[0].cnt)); })
        .catch(function(){});
      runQuery(readForm());
    })
    .catch(function(e){
      list.innerHTML='<div class="state error"><div class="ico">'+
        '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9 2.4 18a2 2 0 0 0 1.7 3h15.8a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/></svg>'+
        '</div><h3>No fue posible conectar con el origen de datos</h3>'+
        '<p>'+escHtml((e&&e.message)||"Error de red")+'. Para que el navegador permita la consulta a datos.gov.co, sirve la página por <b>https</b> (GitHub Pages ya lo hace).</p></div>';
      rcount.textContent="Sin conexión";
    });
  })();
})();
