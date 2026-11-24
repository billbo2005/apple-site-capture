// Project Apple Site Capture – settings (none of these are secrets)
window.APP_CONFIG = {
  clientId: "0778014d-fa03-4f60-907c-c2be61102b57",       // Entra app registration "Apple Site Capture"
  tenantId: "0626ec6c-a8af-4237-aced-7b531aa4987b",       // Project Central (IPC) tenant
  siteHost: "projectsupport.sharepoint.com",
  sitePath: "/sites/AppleSiteCapture2",
  itemsList: "WorksToGo",
  roomsList: "Rooms",
  photosList: "Photos",
  newIdsFrom: 30001,                                     // items raised in the app are numbered from WTG-30001 (the tracker's rows, spare ones included, run to WTG-20400)
  photoToClose: false,                                   // true = a photo is needed before an activity can be closed (as in v1.0)
  version: "1.2.0"
};
