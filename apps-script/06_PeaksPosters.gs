/**
 * Trail Journal -- Apps Script backend
 * File 6 of N: Mountaineer Peaks poster overrides.
 *
 * The 7 location-specific posters (plus the all-locations summary, used as the Classroom /
 * fallback default) ship as static images in the frontend repo (assets/peaks/*.png), pulled
 * straight from the EMS-Mountaineer-Peaks.pptx deck. This file lets Matt replace any one of
 * them from the admin dashboard -- e.g. when EMS issues an updated poster design -- WITHOUT
 * a code deploy. Uploads go to a dedicated Drive folder and the resulting file's direct-view
 * URL is written to site_config as 'peaks_poster_override_<location>'. The student-facing
 * page (index.html) reads all site_config rows on load and prefers an override URL over its
 * bundled assets/peaks/ default whenever one is set.
 */

// Keys must match the location strings used everywhere else (q-where dropdown,
// LOCATION_PEAKS_DISPLAY_/PEAKS_POSTER_IMAGE_ in index.html, LOCATION_PEAKS_ in
// 02_AIProxy.gs), plus the one synthetic 'Summary' key for the Classroom/Other fallback
// poster (PEAKS_POSTER_FALLBACK_ in index.html).
const PEAKS_POSTER_LOCATIONS_ = [
  'Hallways', 'Cafeteria', 'Auditorium', 'Bathroom', 'Media Center',
  'Traveling To & From School', 'Digital Environment', 'Summary',
];

function peaksPosterConfigKey_(location) {
  return 'peaks_poster_override_' + location;
}

/**
 * payload: { location, filename, mimeType, base64 } -- base64 is the raw file content with
 * no "data:image/png;base64," prefix (the admin dashboard strips that before sending).
 * Gated admin-only -- see the site_config write check in enforceAdminGate_ (05_AdminAuth.gs).
 */
function uploadPeaksPoster_(payload) {
  const location = payload && payload.location;
  if (PEAKS_POSTER_LOCATIONS_.indexOf(location) === -1) {
    throw new Error('Unknown Peaks poster location: ' + location);
  }
  if (!payload.base64) throw new Error('No image data received.');

  const folder = getOrCreatePeaksPosterFolder_();
  const bytes = Utilities.base64Decode(payload.base64);
  const blob = Utilities.newBlob(bytes, payload.mimeType || 'image/png', payload.filename || (location + '.png'));

  // Replace, don't accumulate: trash any previous upload for this exact location first.
  const prefix = 'peaks-poster -- ' + location + ' -- ';
  const existing = folder.getFiles();
  while (existing.hasNext()) {
    const f = existing.next();
    if (f.getName().indexOf(prefix) === 0) f.setTrashed(true);
  }

  const file = folder.createFile(blob).setName(prefix + (payload.filename || 'poster.png'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600';

  setSiteConfigValue_(peaksPosterConfigKey_(location), url);
  return { ok: true, url: url, location: location };
}

/** Lets the admin dashboard clear an override and fall back to the bundled default. */
function resetPeaksPoster_(payload) {
  const location = payload && payload.location;
  if (PEAKS_POSTER_LOCATIONS_.indexOf(location) === -1) {
    throw new Error('Unknown Peaks poster location: ' + location);
  }
  deleteSiteConfigValue_(peaksPosterConfigKey_(location));
  return { ok: true, location: location };
}

function getOrCreatePeaksPosterFolder_() {
  const name = 'Trail Journal -- Mountaineer Peaks Posters';
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}
