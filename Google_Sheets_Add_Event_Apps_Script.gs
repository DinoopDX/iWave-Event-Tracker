/** iWave Event Dashboard - public Add Event form endpoint.
 * Paste this into Apps Script attached to the target Google Sheet.
 * Deploy as Web App: Execute as Me, Access: Anyone.
 */
const SPREADSHEET_ID = '12F6rJ1C5wRYuuHjSfDSLx6W4LykCKLUn_KggggykjNA';
const EVENTS_TAB_NAME = 'Events';


function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!String(payload.name || '').trim() || !String(payload.region || '').trim()) {
      return json_({ok:false, error:'Event name and region/city are required.'});
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(EVENTS_TAB_NAME);
    if (!sheet) throw new Error('Events sheet not found');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const data = {
      event: payload.name || '',
      month: payload.month || '',
      dates: payload.dates || '',
      start: payload.start || '',
      end: payload.end || '',
      region: payload.region || '',
      category: payload.category || '',
      attendees: Array.isArray(payload.attendees) ? payload.attendees.join(', ') : (payload.attendees || ''),
      leads: payload.leads || '', contacts: payload.contacts || '',
      reg: payload.reg || '', setup: payload.setup || '', mkt: payload.mkt || '',
      logistics: payload.logistics || '', travelCost: payload.travelCost || '', booth: payload.booth || '',
      rev3: payload.rev3 || '', rev6: payload.rev6 || '', rev12: payload.rev12 || '',
      feedback: payload.feedback || '', notes: payload.notes || ''
    };
    const aliases = {
      event:['event','eventname','name'], month:['month'], dates:['datesasshown','dates'],
      start:['startdate','start'], end:['enddate','end'], region:['regioncity','region','city'],
      category:['category'], attendees:['attendeescomma separated','attendees','attendeescomma separated'],
      leads:['leads'], contacts:['contacts'], reg:['registrationusd','registration'], setup:['boothsetupusd','boothsetup'],
      mkt:['marketingusd','marketing'], logistics:['logisticsusd','logistics'], travelCost:['travelstayusd','travelandstayusd','travelstay'],
      booth:['boothno','boothnumber'], rev3:['revenueexpected3mousd','revenueexpected3monthsusd','revenueexpected3mo'],
      rev6:['revenueexpected6mousd','revenueexpected6monthsusd','revenueexpected6mo'], rev12:['revenueexpected1yrusd','revenueexpected1yearusd'],
      feedback:['feedback'], notes:['notes']
    };
    const row = headers.map(h => {
      const hn = norm(h);
      for (const key in aliases) if (aliases[key].some(a => hn === norm(a) || hn.includes(norm(a)))) return data[key] || '';
      return '';
    });
    sheet.appendRow(row);
    return json_({ok:true, message:'Event added'});
  } catch (err) {
    return json_({ok:false, error:String(err)});
  }
}
function doGet() { return json_({ok:true, message:'iWave Event form endpoint is running'}); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
