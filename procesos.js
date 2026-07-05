"use strict";
(function(){
  /* ===== configuración ===== */
  /* API SODA de Socrata (datos.gov.co): soporta CORS desde el navegador.
     Dataset: SECOP II - Procesos de Contratación (p6dx-8zbt).
     (El dataset 77td-mmia está vacío, por eso no se usa.) */
  var API = "https://www.datos.gov.co/resource/p6dx-8zbt.json";
  var PAGE = 20;

  /* documentos de proveedor restringidos (igual que en Contratos) */
  var BLOCKED_DOCS = ["1128272022"];

  /* nombres reales de campos del dataset p6dx-8zbt (con respaldos por si cambian) */
  var FIELD_CANDIDATES = {
    nomEnt:      ["entidad"],                                  /* nombre de la entidad (texto) */
    nitEnt:      ["nit_entidad"],                              /* NIT entidad (texto) */
    desc:        ["descripci_n_del_procedimiento","descripcion_del_procedimiento"],
    referencia:  ["referencia_del_proceso","referencia"],
    nitProv:     ["nit_del_proveedor_adjudicado","nit_del_proveedor"],   /* texto */
    nomProv:     ["nombre_del_proveedor","nombre_del_adjudicador"],
    valor:       ["valor_total_adjudicacion","precio_base"],   /* número */
    modalidad:   ["modalidad_de_contratacion","modalidad"],
    objeto:      ["nombre_del_procedimiento","objeto_del_contrato","objeto"],
    fecha:       ["fecha_de_publicacion_del","fecha_de_publicacion_fase_3","fecha_de_ultima_publicaci","fecha_de_publicacion"],
    estado:      ["estado_del_procedimiento","estado_resumen","fase","estado_de_apertura_del_proceso"],
    url:         ["urlproceso","url_del_proceso","url_proceso","url"],
    departamento:["departamento_entidad","departamento"],
    ciudad:      ["ciudad_entidad","ciudad","municipio"]
  };

  function modTheme(m){
    var s=(m||"").toLowerCase();
    if(s.indexOf("licitaci")>=0) return ["#0E8A64","#DCF2E9","#0A5C44"];
    if(s.indexOf("mínima")>=0||s.indexOf("minima")>=0) return ["#B9862B","#F6ECD4","#7A5C16"];
    if(s.indexOf("abreviada")>=0) return ["#7A4FB0","#EEE6F6","#553584"];
    if(s.indexOf("méritos")>=0||s.indexOf("meritos")>=0) return ["#0E6E78","#DFF0F2","#0A555D"];
    if(s.indexOf("subasta")>=0) return ["#3457A6","#E4EAF7","#243F78"];
    if(s.indexOf("especial")>=0||s.indexOf("regimen")>=0||s.indexOf("régimen")>=0) return ["#B4452F","#F6E1DB","#822F1F"];
    return ["#33413A","#E8ECE9","#1C2621"];
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
  /* publicado hace 7 días o menos => insignia "Nuevo" en la tarjeta */
  var NEW_MS=7*864e5;
  function isNuevo(v){ if(!v) return false; var d=new Date(v); if(isNaN(d)) return false;
    var dif=Date.now()-d.getTime(); return dif<=NEW_MS && dif>=-864e5; }
  /* consulta por defecto: licitaciones publicadas en el último mes.
     La marca de tiempo se fija UNA vez por consulta (recentSince) para que la
     paginación, el conteo y la descarga a Excel usen exactamente el mismo corte. */
  function recentQuery(){
    var d=new Date(); d.setMonth(d.getMonth()-1);
    return { recent:true, recentSince:d.toISOString().slice(0,19) };
  }

  /* fetch con timeout: evita que una conexión colgada deje la UI en "Buscando…"
     para siempre. Sin coste de rendimiento; solo aborta peticiones sin respuesta. */
  var FETCH_TIMEOUT=15000;
  function fetchT(url){
    if(typeof AbortController==="undefined") return fetch(url);
    var ac=new AbortController(), t=setTimeout(function(){ ac.abort(); }, FETCH_TIMEOUT);
    return fetch(url,{signal:ac.signal}).then(
      function(r){ clearTimeout(t); return r; },
      function(e){ clearTimeout(t); throw e; });
  }

  /* ===== estado ===== */
  var F=null, active=null, rows=[], count=null, page=0,
      loading=false, more=false, moreErr=false, done=false, error=null, reqId=0, xlsBusy=false;
  var $=function(id){ return document.getElementById(id); };
  var list=$("list"), rcount=$("rcount"), rsub=$("rsub");

  /* poblar el desplegable de años */
  (function(){
    var y=$("f_anio"), now=new Date().getFullYear();
    for(var yr=now; yr>=2015; yr--){ var o=document.createElement("option"); o.value=yr;o.textContent=yr;y.appendChild(o); }
  })();

  /* ===== consultas SODA (SoQL) ===== */
  /* En este dataset: NIT entidad y NIT proveedor son TEXTO (coincidencia
     exacta con comillas); valor_total_adjudicacion es NÚMERO (comparación
     numérica). Los campos de texto libre usan lower()+like (contiene). */
  /* El aplicativo excluye SIEMPRE la contratación directa. En el respaldo
     sin lower(), '%irecta%' cubre "Directa" y "directa". */
  function noDirecta(opts){
    return "not ("+((opts&&opts.lower)
      ? "lower("+F.modalidad+") like '%directa%'"
      : F.modalidad+" like '%irecta%'")+")";
  }
  function buildWhere(a, opts){
    if(!F) return "";
    var p=[];
    function txt(field, value){                 /* texto libre: contiene */
      if(!value || !field) return;
      var v=esc(opts.lower? value.toLowerCase() : value);
      var lhs=opts.lower? "lower("+field+")" : field;
      if(opts.escW){                            /* trata %, _ y \ del usuario como literales */
        p.push(lhs+" like '%"+v.replace(/[\\%_]/g,"\\$&")+"%' escape '\\'");
      } else {                                  /* respaldo: si el servidor no soporta ESCAPE */
        p.push(lhs+" like '%"+v+"%'");
      }
    }
    function idTxt(field, value){                /* identificador de texto: exacto */
      if(!value || !field) return;
      p.push(field+" = '"+esc(String(value).trim())+"'");
    }
    function gte(field, value){                  /* numérico: valor mínimo */
      if(!value || !field) return;
      var raw=String(value).trim();
      if(/^[0-9]+$/.test(raw)) p.push(field+" >= "+raw);
    }
    idTxt(F.nitEnt,    a.nitEnt);
    txt(F.nomEnt,      a.nomEnt);
    txt(F.desc,        a.desc);
    txt(F.referencia,  a.ref);
    idTxt(F.nitProv,   a.nitProv);
    txt(F.nomProv,     a.nomProv);
    gte(F.valor,       a.valorMin);
    txt(F.modalidad,   a.mod);
    txt(F.objeto,      a.objeto);
    if(a.anio && F.fecha){ var yy=parseInt(a.anio,10);
      p.push(F.fecha+" >= '"+yy+"-01-01T00:00:00' and "+F.fecha+" < '"+(yy+1)+"-01-01T00:00:00'"); }
    if(a.recent && a.recentSince && F.fecha){   /* vista por defecto: último mes */
      p.push(F.fecha+" >= '"+esc(a.recentSince)+"'"); }
    if(F.modalidad){ p.push(noDirecta(opts)); } /* TODA consulta excluye la contratación directa */
    if(F.nitProv){                               /* excluir proveedores restringidos */
      BLOCKED_DOCS.forEach(function(doc){ p.push(F.nitProv+" != '"+esc(doc)+"'"); });
    }
    return p.join(" and ");
  }
  /* orden: por fecha (más reciente primero) y los procesos sin fecha al final */
  /* :id (columna de sistema de Socrata) como desempate => paginación estable:
     evita filas duplicadas/omitidas entre páginas/lotes cuando hay empates de
     fecha. Se dejan respaldos SIN :id por si el dataset no expone esa columna. */
  function orderCandidates(){
    if(!(F && F.fecha)) return [":id", ""];
    var f=F.fecha;
    return ["coalesce("+f+",'1111-01-01T00:00:00') desc, :id", f+" desc, :id",
            "coalesce("+f+",'1111-01-01T00:00:00') desc", f+" desc"];
  }
  /* Estrategias de respaldo ante un 400, en CASCADA (no producto cartesiano: así
     el peor caso son pocas peticiones, no decenas). Se degrada por pasos —primero
     quita ESCAPE, luego :id/coalesce, luego lower()— y la ÚLTIMA combinación
     reproduce el comportamiento original, de modo que ningún fix puede romper la
     búsqueda. En el camino feliz gana la 1ª (una sola petición, count en paralelo). */
  function buildStrategies(){
    var ords=orderCandidates(), s=[];
    s.push({ord:ords[0], opts:{lower:true, escW:true}});     /* orden + texto más ricos */
    s.push({ord:ords[0], opts:{lower:true, escW:false}});    /* sin ESCAPE */
    for(var i=1;i<ords.length;i++)                            /* degradar el orden */
      s.push({ord:ords[i], opts:{lower:true, escW:false}});
    s.push({ord:ords[ords.length-1], opts:{lower:false, escW:false}});  /* == original */
    return s;
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
      var pData=fetchT(dataUrl).then(function(r){ if(!r.ok) throw new Error("HTTP "+r.status); return r.json(); });
      if(!withCount) return pData.then(function(arr){ return {value:arr}; });
      var cntUrl=API+"?$select="+encodeURIComponent("count(1) as cnt")+(w? "&$where="+encodeURIComponent(w):"");
      var pCount=fetchT(cntUrl).then(function(r){ return r.ok?r.json():null; }).catch(function(){ return null; });
      return Promise.all([pData,pCount]).then(function(res){
        var out={value:res[0]};
        if(res[1] && res[1][0] && res[1][0].cnt!=null) out["@odata.count"]=Number(res[1][0].cnt);
        return out;
      });
    }
    return tryStrategies(attempt);
  }

  /* ===== descarga a Excel ===== */
  var XLS_CAP=5000, XLS_BATCH=1000;
  var XLS_COLS=[
    ["nomEnt","Entidad"], ["nitEnt","NIT entidad"],
    ["objeto","Objeto"], ["desc","Descripción del procedimiento"],
    ["referencia","Referencia del proceso"], ["modalidad","Modalidad"],
    ["nomProv","Proveedor adjudicado"], ["nitProv","NIT proveedor"],
    ["valor","Valor total adjudicado"], ["fecha","Fecha de publicación"],
    ["estado","Estado"], ["departamento","Departamento"],
    ["ciudad","Ciudad"], ["url","URL del proceso"]
  ];
  function fetchRows(a, offset, limit){
    function attempt(st){
      var w=buildWhere(a,st.opts);
      var url=API+"?$limit="+limit+"&$offset="+offset;
      if(w) url+="&$where="+encodeURIComponent(w);
      if(st.ord) url+="&$order="+encodeURIComponent(st.ord);
      return fetchT(url).then(function(r){ if(!r.ok) throw new Error("HTTP "+r.status); return r.json(); });
    }
    return tryStrategies(attempt);
  }
  function buildRecords(allRows){
    return allRows.map(function(row){
      var rec={};
      XLS_COLS.forEach(function(c){
        var v=val(row, F[c[0]]);
        if(c[0]==="valor"){ var n=(v!=null&&v!=="")?Number(v):null; rec[c[1]]=(n!=null&&!isNaN(n))?n:(v||""); }
        else if(c[0]==="fecha"){ rec[c[1]]=v?String(v).slice(0,10):""; }
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
  /* da formato al .xlsx: anchos de columna, encabezado con estilo, filtros y miles */
  function styleSheet(ws){
    var range=XLSX.utils.decode_range(ws['!ref']);
    var c, r, cell, head, maxw=[], isValor=[];
    /* encabezados y detección de columnas "Valor" (una vez por columna) */
    for(c=range.s.c;c<=range.e.c;c++){
      maxw[c]=10;
      head=ws[XLSX.utils.encode_cell({c:c,r:0})];
      isValor[c]= !!(head && /^Valor/i.test(String(head.v)));
      if(head){
        head.s={ font:{bold:true,color:{rgb:"FFFFFF"}}, fill:{fgColor:{rgb:"0A1A2F"}},
          alignment:{horizontal:"center",vertical:"center",wrapText:true},
          border:{bottom:{style:"medium",color:{rgb:"B98A28"}}} };
      }
    }
    /* un solo recorrido de celdas: ancho de columna + formato de miles */
    for(r=range.s.r;r<=range.e.r;r++){
      for(c=range.s.c;c<=range.e.c;c++){
        cell=ws[XLSX.utils.encode_cell({c:c,r:r})];
        if(!cell||cell.v==null) continue;
        var len=String(cell.v).length; if(len>maxw[c]) maxw[c]=len;
        if(r>0 && isValor[c] && typeof cell.v==="number") cell.z="#,##0";
      }
    }
    var cols=[]; for(c=range.s.c;c<=range.e.c;c++) cols.push({wch:Math.min(maxw[c]+2,60)});
    ws['!cols']=cols;
    ws['!autofilter']={ref:ws['!ref']};
  }
  function exportRecords(records, base){
    if(window.XLSX){
      var ws=XLSX.utils.json_to_sheet(records);
      styleSheet(ws);
      var wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Procesos");
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
    var lbl=$("xlsLabel"); if(lbl) lbl.textContent = on? " Generando…" : " Descargar Excel";
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
      if(!all.length){ alert("No hay procesos para descargar con los filtros actuales."); return; }
      var base=(a&&a.recent)? "licitaciones-ultimo-mes-" : "procesos-secop-ii-";
      exportRecords(buildRecords(all), base+new Date().toISOString().slice(0,10));
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
    var ent=val(row,F.nomEnt), nitE=val(row,F.nitEnt), prov=val(row,F.nomProv),
        docP=val(row,F.nitProv), obj=val(row,F.objeto)||val(row,F.desc),
        ref=val(row,F.referencia), mod=val(row,F.modalidad), valor=val(row,F.valor),
        fecha=val(row,F.fecha), url=val(row,F.url), depto=val(row,F.departamento),
        ciudad=val(row,F.ciudad), estado=val(row,F.estado);

    var th=modTheme(mod), accent=th[0], soft=th[1], aink=th[2];
    var year=getYear(fecha);
    var valNum=(valor!=null && valor!=="")? Number(valor):null;
    var loc=[ciudad,depto].filter(Boolean).join(", ");
    var provDef=(prov && prov!=="No Definido"), docDef=(docP && docP!=="No Definido");

    var top='<div class="card-top">';
    if(mod)            top+='<span class="tag mod">'+escHtml(mod)+'</span>';
    if(isNuevo(fecha)) top+='<span class="tag new">Nuevo</span>';
    top+='<span class="when">'+(fecha? 'Publicado · '+escHtml(fmtFecha(fecha)) : 'Sin fecha')+'</span></div>';

    var meta="";
    if(nitE)   meta+='<span><span class="k">NIT</span> '+escHtml(nitE)+'</span>';
    if(loc)    meta+='<span>'+escHtml(loc)+'</span>';
    if(ref)    meta+='<span><span class="k">Ref.</span> '+escHtml(ref)+'</span>';
    if(estado) meta+='<span><span class="k">Estado</span> '+escHtml(estado)+'</span>';

    var provHtml="";
    if(provDef){ provHtml='<div class="prov"><span class="plabel">Proveedor adjudicado</span>'+
      '<span class="who">'+escHtml(prov)+'</span>'+(docDef? '<span class="nit">· '+escHtml(docP)+'</span>':'')+'</div>'; }

    var safeUrl=(url && /^https?:\/\//i.test(url))? url : null;
    var verproc = safeUrl
      ? '<a class="verproc" href="'+escHtml(safeUrl)+'" target="_blank" rel="noopener noreferrer">Ver proceso '+
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M7 17 17 7M9 7h8v8"/></svg></a>'
      : '<span class="verproc disabled">Sin enlace</span>';

    return ''+
    '<article class="card" style="--accent:'+accent+';--accent-soft:'+soft+';--accent-ink:'+aink+'">'+
      top+
      '<h2 class="ent">'+escHtml(ent||"Entidad no registrada")+'</h2>'+
      (meta? '<div class="meta">'+meta+'</div>':'')+
      (obj? '<p class="objeto">'+escHtml(obj)+'</p>':'')+
      provHtml+
      '<div class="card-foot">'+
        '<div class="money">'+
          '<span class="vlbl">Valor total adjudicado</span>'+
          '<span class="valor">'+((valNum!=null && !isNaN(valNum))? COP.format(valNum):'—')+'</span>'+
          (year? '<span class="anio">Año '+year+'</span>':'')+
        '</div>'+ verproc+
      '</div>'+
    '</article>';
  }

  function render(){
    var isRecent = !!(active && active.recent);
    var badge=$("viewbadge"); if(badge) badge.hidden=!isRecent;
    if(loading){ rcount.innerHTML="Buscando…"; }
    else if(count!=null){
      rcount.innerHTML = isRecent
        ? '<span>'+NUM.format(count)+'</span> '+(count===1?'licitación publicada':'licitaciones publicadas')+' en el último mes'
        : '<span>'+NUM.format(count)+'</span> proceso'+(count===1?'':'s')+' encontrado'+(count===1?'':'s');
    }
    else { rcount.textContent=rows.length+" resultado"+(rows.length===1?'':'s'); }
    var hasF = active && (active.nitEnt||active.nomEnt||active.desc||active.ref||active.nitProv||active.nomProv||active.valorMin||active.mod||active.objeto||active.anio);
    rsub.textContent = isRecent? "Sin contratación directa · las más recientes primero · desde "+(fmtFecha(active.recentSince)||"hace un mes")
      : (hasF? "Según los filtros aplicados · sin contratación directa" : "Muestra del registro nacional · sin contratación directa");

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
        (isRecent
          ? '</div><h3>Sin licitaciones recientes</h3>'+
            '<p>El conjunto de datos aún no registra procesos publicados en el último mes. Usa los filtros para consultar el histórico.</p></div>'
          : '</div><h3>Sin coincidencias</h3>'+
            '<p>Ningún proceso coincide con esos filtros. Prueba términos más generales o limpia alguno.</p></div>');
      return;
    }
    for(var j=0;j<rows.length;j++) html+=cardHtml(rows[j]);
    if(!done) html+=loadMoreHtml();
    list.innerHTML=html;
    bindLoadMore();
  }
  function loadMoreHtml(){
    return '<div class="loadmore"><button class="btn" id="loadmore"'+(more?' disabled':'')+'>'+
      (more? '<span class="spin dark"></span>Cargando…'
           : (moreErr? 'Error al cargar · Reintentar' : 'Cargar más resultados'))+'</button></div>';
  }
  function bindLoadMore(){ var lm=$("loadmore"); if(lm) lm.addEventListener("click", loadMore); }
  /* repinta solo el botón "Cargar más" sin reconstruir las tarjetas ya visibles */
  function syncMore(){
    var wrap=list.querySelector(".loadmore"); if(wrap) wrap.parentNode.removeChild(wrap);
    if(!done){ list.insertAdjacentHTML("beforeend", loadMoreHtml()); bindLoadMore(); }
  }
  /* anexa solo las tarjetas nuevas (evita el re-render O(n) en cada página) */
  function appendCards(newRows){
    var h=""; for(var j=0;j<newRows.length;j++) h+=cardHtml(newRows[j]);
    var wrap=list.querySelector(".loadmore");
    if(wrap) wrap.insertAdjacentHTML("beforebegin", h);
    else list.insertAdjacentHTML("beforeend", h);
  }

  /* ===== acciones ===== */
  function runQuery(a){
    active=a; var id=++reqId;
    loading=true; error=null; rows=[]; count=null; page=0; done=false; more=false; moreErr=false;
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
    var id=reqId, next=page+1; more=true; moreErr=false; syncMore();
    fetchPage(active,next,false).then(function(j){
      if(id!==reqId) return;
      var v=j.value||[]; rows=rows.concat(v); page=next; done=v.length<PAGE;
      appendCards(v);   /* conserva las tarjetas ya cargadas; solo añade las nuevas */
    }).catch(function(e){ if(id===reqId) moreErr=true; })   /* error de página: no borra lo ya mostrado */
    .then(function(){ if(id===reqId){ more=false; syncMore(); } });
  }
  function readForm(){
    return {
      nitEnt:$("f_nitEnt").value.trim(), nomEnt:$("f_nomEnt").value.trim(),
      desc:$("f_desc").value.trim(), ref:$("f_ref").value.trim(), anio:$("f_anio").value,
      nitProv:$("f_nitProv").value.trim(), nomProv:$("f_nomProv").value.trim(),
      valorMin:$("f_valor").value.trim(), mod:$("f_mod").value.trim(), objeto:$("f_objeto").value.trim()
    };
  }
  var INPUT_IDS=["f_nomEnt","f_nitEnt","f_desc","f_ref","f_anio","f_nitProv","f_nomProv","f_valor","f_mod","f_objeto"];

  /* vista por defecto: últimas licitaciones publicadas (último mes) */
  function showRecent(){ runQuery(recentQuery()); }
  function showBlocked(){
    active=null; rows=[]; count=null; error=null; loading=false; done=true;
    setBtnLoading(false);
    var badge=$("viewbadge"); if(badge) badge.hidden=true;
    rcount.textContent="Búsqueda no permitida";
    rsub.textContent="Este proveedor está restringido";
    list.innerHTML='<div class="state"><div class="ico">'+
      '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>'+
      '</div><h3>Proveedor restringido</h3>'+
      '<p>No es posible consultar procesos por este documento de proveedor.</p></div>';
  }
  function isBlocked(a){ return a.nitProv && BLOCKED_DOCS.indexOf(a.nitProv)>=0; }

  function hasFilters(a){
    return !!(a.nitEnt||a.nomEnt||a.desc||a.ref||a.nitProv||a.nomProv||a.valorMin||a.mod||a.objeto||a.anio);
  }
  $("form").addEventListener("submit", function(e){ e.preventDefault(); if(!F) return;
    var a=readForm(); if(isBlocked(a)){ showBlocked(); return; }
    if(!hasFilters(a)){ showRecent(); return; }   /* sin filtros => últimas licitaciones */
    runQuery(a); });
  $("clear").addEventListener("click", function(){
    INPUT_IDS.forEach(function(id){ $(id).value=""; });
    if(F) showRecent();
  });
  $("btnXlsx").addEventListener("click", downloadExcel);

  /* campos numéricos: solo dígitos (escritura y pegado) */
  ["f_nitEnt","f_nitProv","f_valor"].forEach(function(id){
    $(id).addEventListener("input", function(){
      var clean=this.value.replace(/\D+/g,"");
      if(this.value!==clean) this.value=clean;
    });
  });

  /* ===== arranque ===== */
  (function boot(){
    /* Socrata omite del JSON los campos nulos de cada fila, así que mapear el
       esquema sobre UNA sola fila puede dejar campos sin resolver. Se unen las
       claves de varias filas para detectar el esquema de forma fiable. El coste
       es una única petición de arranque algo mayor, no afecta a las búsquedas. */
    fetchT(API+"?$limit=25").then(function(r){ if(!r.ok) throw new Error("HTTP "+r.status); return r.json(); })
    .then(function(j){
      var arr=(j&&j.length)?j:[], keySet={};
      arr.forEach(function(row){ if(row&&typeof row==="object") for(var kk in row) keySet[kk]=1; });
      var keys=Object.keys(keySet);
      if(!keys.length){   /* dataset vacío/no disponible: no habilitar búsquedas que ignorarían los filtros */
        list.innerHTML='<div class="state error"><div class="ico">'+
          '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9 2.4 18a2 2 0 0 0 1.7 3h15.8a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/></svg>'+
          '</div><h3>El conjunto de datos no está disponible</h3>'+
          '<p>El origen respondió sin registros. Inténtalo de nuevo en unos minutos.</p></div>';
        rcount.textContent="Sin datos"; return;
      }
      F={};
      for(var k in FIELD_CANDIDATES) F[k]=pick(keys,FIELD_CANDIDATES[k]);
      setBtnLoading(false); setXlsBusy(false);
      /* contador del encabezado: también sin contratación directa (con
         respaldo al conteo simple si el servidor rechaza la cláusula) */
      var cntUrl=API+"?$select="+encodeURIComponent("count(1) as cnt");
      (F.modalidad
        ? fetchT(cntUrl+"&$where="+encodeURIComponent(noDirecta({lower:true})))
            .then(function(x){ if(!x.ok) throw new Error("HTTP "+x.status); return x.json(); })
            .catch(function(){ return fetchT(cntUrl).then(function(x){ return x.ok?x.json():null; }); })
        : fetchT(cntUrl).then(function(x){ return x.ok?x.json():null; })
      ).then(function(d){ if(d && d[0] && d[0].cnt!=null) $("total").textContent=NUM.format(Number(d[0].cnt)); })
       .catch(function(){});
      showRecent();
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
