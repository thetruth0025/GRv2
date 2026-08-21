'use strict';

const { requestJson, HttpError } = require('./http');
const { parseMoney, parseQuantity } = require('./normalize');

const SUPPLIER = 'Mouser';
const BASE = 'https://api.mouser.com/api/v1';

// Mouser authenticates with a single API key on the query string and reports
// errors inside a 200 response, so success has to be checked in the body.
class MouserClient {
  constructor(options) {
    const opts = options || {};
    this.apiKey = opts.apiKey;
    this.currency = opts.currency || 'USD';
  }

  get id() {
    return 'mouser';
  }

  get name() {
    return SUPPLIER;
  }

  get configured() {
    return !!this.apiKey;
  }

  async search(keyword, records) {
    const url = BASE + '/search/keyword?apiKey=' + encodeURIComponent(this.apiKey);
    const { data } = await requestJson(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify({
        SearchByKeywordRequest: {
          keyword,
          records: records || 10,
          startingRecord: 0,
          searchOptions: '',
          searchWithYourSignUpLanguage: '',
        },
      }),
    });

    if (data && Array.isArray(data.Errors) && data.Errors.length > 0) {
      const message = data.Errors.map((e) => e.Message || e.Code).filter(Boolean).join('; ');
      throw new HttpError('Mouser API error: ' + (message || 'unknown error'), 400, data.Errors);
    }
    return data || {};
  }

  toRecord(data, part) {
    const results = (data && data.SearchResults) || {};
    const parts = Array.isArray(results.Parts) ? results.Parts : [];
    if (parts.length === 0) return null;
    const match = pickBestPart(parts, part.mpn, part.manufacturer);
    if (!match) return null;
    return buildRecord(match, part, this.currency, parts.length);
  }

  async fetchRecord(part) {
    return this.toRecord(await this.search(part.mpn, 10), part);
  }
}

function normalizeKey(text) {
  return String(text || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function pickBestPart(parts, mpn, manufacturer) {
  const wantMpn = normalizeKey(mpn);
  if (!wantMpn) return null;
  const wantMfr = normalizeKey(manufacturer);

  let best = null;
  let bestScore = -1;
  for (const candidate of parts) {
    const candidateMpn = normalizeKey(candidate.ManufacturerPartNumber);
    const candidateMfr = normalizeKey(candidate.Manufacturer);
    let score = 0;
    if (candidateMpn && candidateMpn === wantMpn) score += 100;
    else if (candidateMpn && (candidateMpn.startsWith(wantMpn) || wantMpn.startsWith(candidateMpn))) score += 50;
    else if (candidateMpn && (candidateMpn.includes(wantMpn) || wantMpn.includes(candidateMpn))) score += 20;
    if (wantMfr && candidateMfr && (candidateMfr === wantMfr || candidateMfr.includes(wantMfr) || wantMfr.includes(candidateMfr))) {
      score += 30;
    }
    if (stockOf(candidate) > 0) score += 5;
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }
  return bestScore >= 20 ? best : null;
}

// Availability arrives as prose ("1,234 In Stock", "None"); AvailabilityInStock
// is the clean number when Mouser includes it.
function stockOf(part) {
  const direct = parseQuantity(part.AvailabilityInStock);
  if (Number.isFinite(direct)) return direct;
  const availability = String(part.Availability || '').trim();
  if (!availability) return 0;
  if (/^none$/i.test(availability)) return 0;
  const parsed = parseQuantity(availability);
  return Number.isFinite(parsed) ? parsed : 0;
}

function priceBreaksOf(part) {
  const list = Array.isArray(part.PriceBreaks) ? part.PriceBreaks : [];
  return list
    .map((b) => ({
      quantity: parseQuantity(b.Quantity),
      unitPrice: parseMoney(b.Price),
      currency: b.Currency || null,
    }))
    .filter((b) => Number.isFinite(b.quantity) && Number.isFinite(b.unitPrice))
    .sort((a, b) => a.quantity - b.quantity);
}

function lifecycleOf(part) {
  // LifecycleStatus is authoritative but frequently blank; ProductStatus is the
  // Mouser-catalog view ("New at Mouser", "Obsolete") and fills the gap.
  const lifecycle = String(part.LifecycleStatus || '').trim();
  if (lifecycle) return lifecycle;
  const status = String(part.ProductStatus || '').trim();
  if (status) return status;
  return null;
}

function buildRecord(part, bomPart, currency, matchCount) {
  const priceBreaks = priceBreaksOf(part);
  const stock = stockOf(part);
  const record = {
    supplier: SUPPLIER,
    manufacturer: part.Manufacturer || null,
    manufacturerPartNumber: part.ManufacturerPartNumber || null,
    description: part.Description || null,
    productUrl: part.ProductDetailUrl || null,
    datasheetUrl: part.DataSheetUrl || null,
    leadTime: part.LeadTime || null,
    lifecycle: lifecycleOf(part),
    rohs: part.ROHSStatus || null,
    totalStock: stock,
    currency: (priceBreaks[0] && priceBreaks[0].currency) || currency,
    matchCount,
    exactMatch: normalizeKey(part.ManufacturerPartNumber) === normalizeKey(bomPart.mpn),
    factoryStock: parseQuantity(part.FactoryStock),
    // Mouser sells one catalog line per part number, so there is a single
    // packaging option rather than DigiKey's variation list.
    variations: [
      {
        supplierPartNumber: part.MouserPartNumber || null,
        packaging: part.Reeling ? 'Reel' : null,
        stock,
        minimumOrderQuantity: parseQuantity(part.Min) || 1,
        orderMultiple: parseQuantity(part.Mult) || 1,
        priceBreaks,
        marketPlace: false,
      },
    ],
  };

  const replacement = String(part.SuggestedReplacement || '').trim();
  if (replacement) record.suggestedReplacement = replacement;
  return record;
}

module.exports = { MouserClient, SUPPLIER };
