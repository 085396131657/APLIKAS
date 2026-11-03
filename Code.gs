// Perpustakaan Digital Sekolah: backend Google Apps Script + Google Sheets
// GANTI PIN di bawah sebelum deploy. PIN dipakai guru untuk membuka Mode guru.
var PIN = "GANTI-PIN-GURU";

var H = {
  Resensi:   ["id","buku","nama","rating","teks","status","waktu"],
  Usulan:    ["id","judul","alasan","waktu"],
  Aktivitas: ["id","user","tipe","buku","tgl","batas","status"],
  Anggota:   ["id","nama","peran","nomor","kelas","waktu"]
};

function sh(n){
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(n);
  if(!s){ s = ss.insertSheet(n); s.getRange("A:Z").setNumberFormat("@"); s.appendRow(H[n]); }
  return s;
}
function rows(n){ var v = sh(n).getDataRange().getValues(); v.shift(); return v; }
function clean(x){ x = String(x == null ? "" : x).slice(0,1000); return /^[=+\-@]/.test(x) ? " " + x : x; }
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function findRow(n, id){
  var v = sh(n).getDataRange().getValues();
  for(var i=1;i<v.length;i++){ if(String(v[i][0])===String(id)) return i+1; }
  return 0;
}

function doGet(){ return out({ok:true, msg:"Perpustakaan Digital API aktif"}); }

function doPost(e){
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try{
    var q = JSON.parse(e.postData.contents), a = q.a, now = new Date().toISOString();
    var guru = !!q.pin && q.pin === PIN, user = clean(q.user);
    if(a==="rev")  sh("Resensi").appendRow([clean(q.id),clean(q.b),clean(q.n),Number(q.r)||0,clean(q.x),"wait",now]);
    if(a==="sugg") sh("Usulan").appendRow([Date.now(),clean(q.t),clean(q.al),now]);
    if(a==="read") sh("Aktivitas").appendRow([Date.now(),user,"baca",clean(q.b),clean(q.d),"",""]);
    if(a==="loan") sh("Aktivitas").appendRow([clean(q.id),user,"pinjam",clean(q.b),clean(q.d),clean(q.due),"out"]);
    if(a==="ret"){ var r = findRow("Aktivitas", q.id); if(r && sh("Aktivitas").getRange(r,2).getValue()===user) sh("Aktivitas").getRange(r,7).setValue("back"); }
    if(a==="mod" && guru){
      var m = findRow("Resensi", q.id);
      if(m){ if(q.op==="ok") sh("Resensi").getRange(m,6).setValue("ok"); if(q.op==="del") sh("Resensi").deleteRow(m); }
    }
    // khusus guru (butuh PIN): kelola anggota dan catat peminjaman manual
    if(guru && a==="addm") sh("Anggota").appendRow([clean(q.id),clean(q.n),clean(q.p),clean(q.no),clean(q.k),now]);
    if(guru && a==="delm"){ var dm = findRow("Anggota", q.id); if(dm) sh("Anggota").deleteRow(dm); }
    if(guru && a==="gloan") sh("Aktivitas").appendRow([clean(q.id),clean(q.u),"pinjam",clean(q.b),clean(q.d),clean(q.due),"out"]);
    if(guru && a==="gret"){ var gr = findRow("Aktivitas", q.id); if(gr) sh("Aktivitas").getRange(gr,7).setValue("back"); }
    if(a==="clr"){
      var s = sh("Aktivitas"), v = s.getDataRange().getValues();
      for(var i=v.length-1;i>=1;i--){ if(v[i][1]===user && v[i][2]==="baca") s.deleteRow(i+1); }
    }
    // balasan: data terbaru
    var revs = rows("Resensi").filter(function(x){ return x[5]==="ok" || guru; })
      .map(function(x){ return {id:x[0], b:x[1], n:x[2], r:Number(x[3]), x:x[4], s:x[5]}; });
    var sugg = guru ? rows("Usulan").map(function(x){ return {t:x[1], a:x[2], d:x[3]}; }) : [];
    var act = rows("Aktivitas").filter(function(x){ return x[1]===user; });
    var hist = act.filter(function(x){ return x[2]==="baca"; }).map(function(x){ return {b:x[3], d:x[4]}; }).slice(-50);
    var loans = act.filter(function(x){ return x[2]==="pinjam"; })
      .map(function(x){ return {id:x[0], b:x[3], d:x[4], due:x[5], back:x[6]==="back"}; });
    var mem = guru ? rows("Anggota").map(function(x){ return {id:x[0], n:x[1], p:x[2], no:x[3], k:x[4]}; }) : [];
    var all = guru ? rows("Aktivitas").filter(function(x){ return x[2]==="pinjam"; })
      .map(function(x){ return {id:x[0], u:x[1], b:x[3], d:x[4], due:x[5], back:x[6]==="back"}; }) : [];
    return out({ok:true, guru:guru, revs:revs, sugg:sugg, hist:hist, loans:loans, mem:mem, all:all});
  }catch(err){
    return out({ok:false, error:String(err)});
  }finally{
    lock.releaseLock();
  }
}
