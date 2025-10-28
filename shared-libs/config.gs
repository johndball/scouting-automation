// shared-libs/config.gs
var CONFIG_CACHE = null;
function getConfig(){
  if (CONFIG_CACHE) return CONFIG_CACHE;
  var ss = SpreadsheetApp.openById('1AAAAAAAAAAAAAAAAAAAAAAA'); // sample/dummy
  var sh = ss.getSheetByName('Config');
  var map = {};
  var rows = sh.getRange(2,1, sh.getLastRow()-1, 2).getValues();
  rows.forEach(function(r){ var k=r[0], v=r[1]; if (k) map[String(k).trim()] = String(v).trim(); });
  CONFIG_CACHE = map;
  return CONFIG_CACHE;
}
function cfg(key, fallback){ var v = getConfig()[key]; return (v===undefined||v==='') ? fallback : v; }
