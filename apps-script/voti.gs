/**
 * Aglioteca — raccolta voti vetrina.
 * Si incolla in Estensioni > Apps Script del foglio "Aglioteca - Voti vetrina",
 * poi Distribuisci > Nuova distribuzione > App web (Esegui come: me, Accesso: chiunque).
 * L'URL che finisce con /exec va in vetrina.html (costante VOTI_URL).
 */
var MAX = 3;
var CHIAVI = ['abbonamento','scoperta','chilo','ristoranti','nero','regalo','assaggio','vampiri','masterclass'];

function doPost(e) {
  var out = function (o) {
    return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var d = JSON.parse(e.postData.contents);
    if (d.website) return out({ ok: true });                     // trappola anti-bot: finge di andare bene
    var email = String(d.email || '').trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || email.length > 120) return out({ ok: false, error: 'email' });
    var picks = (d.picks || []).filter(function (k) { return CHIAVI.indexOf(k) > -1; });
    picks = picks.filter(function (k, i) { return picks.indexOf(k) === i; });
    if (picks.length < 1 || picks.length > MAX) return out({ ok: false, error: 'picks' });

    var sh = SpreadsheetApp.getActive().getSheetByName('Voti');
    var last = sh.getLastRow();
    if (last > 1) {
      var emails = sh.getRange(2, 2, last - 1, 1).getValues();
      for (var i = 0; i < emails.length; i++) {
        if (String(emails[i][0]).toLowerCase() === email) return out({ ok: false, error: 'duplicate' });
      }
    }
    sh.appendRow([new Date(), email, '|' + picks.join('|') + '|', (d.names || []).slice(0, MAX).join(' / '), d.notify ? 'sì' : 'no']);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}
