'use strict';

const { requestJson, HttpError } = require('./http');
const { parseMoney, parseQuantity } = require('./normalize');

const SUPPLIER = 'DigiKey';
const PROD_BASE = 'https://api.digikey.com';
const SANDBOX_BASE = 'https://sandbox-api.digikey.com';

// DigiKey uses OAuth 2.0 client credentials against Product Information V4.
// Tokens are short-lived (10 minutes in practice), so one is held in memory and
// refreshed just before expiry rather than fetched per part.
class DigiKeyClient {
  constructor(options) {
    const opts = options || {};
    this.clientId = opts.clientId;
    this.clientSecret = opts.clientSecret;
    this.baseUrl = opts.sandbox ? SANDBOX_BASE : PROD_BASE;
    this.sandbox = !!opts.sandbox;
    this.locale = {
      site: opts.site || 'US',
      language: opts.language || 'en',
      currency: opts.currency || 'USD',
    };
    this.token = null;
    this.tokenExpiresAt = 0;
    this.tokenPromise = null;
  }

  get id() {
    return 'digikey';
  }

  get name() {
    return SUPPLIER;
  }

  get configured() {
    return !!(this.clientId && this.clientSecret);
  }

  async getToken() {
    if (this.token && Date.now() < this.tokenExpiresAt) return this.token;
    // Collapse concurrent refreshes so a burst of lookups mints one token.
    if (this.tokenPromise) return this.tokenPromise;

    this.tokenPromise = (async () => {
      const body = new URLSearchParams({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        grant_type: 'client_credentials',
      }).toString();

      const { data } = await requestJson(this.baseUrl + '/v1/oauth2/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body,
        retries: 1,
      });

      if (!data || !data.access_token) {
        throw new HttpError('DigiKey token response did not contain an access_token', 0, data);
      }
      this.token = data.access_token;
      const lifetime = Number(data.expires_in) || 600;
      this.tokenExpiresAt = Date.now() + Math.max(30, lifetime - 60) * 1000;
      return this.token;
    })();

    try {
      return await this.tokenPromise;
    } finally {
      this.tokenPromise = null;
    }
  }

  async search(keyword, limit) {
    const token = await this.getToken();
    const { data } = await requestJson(this.baseUrl + '/products/v4/search/keyword', {
      method: 'POST',
      headers: {
        Authorization: 'Bearer ' + token,
        'X-DIGIKEY-Client-Id': this.clientId,
        'X-DIGIKEY-Locale-Site': this.locale.site,
        'X-DIGIKEY-Locale-Language': this.locale.language,
        'X-DIGIKEY-Locale-Currency': this.locale.currency,
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      body: JSON.stringify({ Keywords: keyword, Limit: limit || 10, Offset: 0 }),
    });
    return data || {};
  }

  // Returns a catalog record, or null when nothing in the response is a
  // credible match for the requested part.
  toRecord(data, part) {
    const products = collectProducts(data);
    if (products.length === 0) return null;
    const product = pickBestProduct(products, part.mpn, part.manufacturer);
    if (!product) return null;
    return buildRecord(product, part, this.locale.currency, products.length);
  }

  async fetchRecord(part) {
    return this.toRecord(await this.search(part.mpn, 10), part);
  }
}

function collectProducts(data) {
  const out = [];
  const seen = new Set();
  const push = (list) => {
    if (!Array.isArray(list)) return;
    for (const item of list) {
      if (!item || typeof item !== 'object') continue;
      const key = mpnOf(item) + '|' + manufacturerOf(item);
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(item);
    }
  };
  // ExactMatches first: V4 puts high-confidence hits there.
  push(data.ExactMatches);
  push(data.Products);
  push(data.ExactManufacturerProducts);
  return out;
}

function mpnOf(product) {
  return String(product.ManufacturerProductNumber || product.ManufacturerPartNumber || '');
}

function manufacturerOf(product) {
  const mfr = product.Manufacturer;
  if (!mfr) return '';
  if (typeof mfr === 'string') return mfr;
  return String(mfr.Name || mfr.Value || '');
}

function descriptionOf(product) {
  const desc = product.Description;
  if (typeof desc === 'string') return desc;
  if (desc && typeof desc === 'object') {
    return desc.ProductDescription || desc.DetailedDescription || null;
  }
  return product.ProductDescription || product.DetailedDescription || null;
}

function statusOf(product) {
  const status = product.ProductStatus;
  if (typeof status === 'string' && status.trim()) return status;
  if (status && typeof status === 'object') {
    const value = status.Status || status.Value;
    if (value) return value;
  }
  // Fall back to the boolean flags V4 also exposes.
  if (product.Obsolete) return 'Obsolete';
  if (product.EndOfLife) return 'End of Life';
  if (product.Discontinued) return 'Discontinued';
  if (product.NormallyStocking) return 'Active';
  return null;
}

function normalizeKey(text) {
  return String(text || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// A keyword search happily returns loosely related parts, so a candidate has to
// clear a similarity floor before it is reported as a match.
function pickBestProduct(products, keyword, manufacturer) {
  const wantMpn = normalizeKey(keyword);
  if (!wantMpn) return null;
  const wantMfr = normalizeKey(manufacturer);

  let best = null;
  let bestScore = -1;
  for (const product of products) {
    const mpn = normalizeKey(mpnOf(product));
    const mfr = normalizeKey(manufacturerOf(product));
    let score = 0;
    if (mpn && mpn === wantMpn) score += 100;
    else if (mpn && (mpn.startsWith(wantMpn) || wantMpn.startsWith(mpn))) score += 50;
    else if (mpn && (mpn.includes(wantMpn) || wantMpn.includes(mpn))) score += 20;
    if (wantMfr && mfr && (mfr === wantMfr || mfr.includes(wantMfr) || wantMfr.includes(mfr))) score += 30;
    if (parseQuantity(product.QuantityAvailable) > 0) score += 5;
    if (score > bestScore) {
      bestScore = score;
      best = product;
    }
  }
  return bestScore >= 20 ? best : null;
}

// V4 nests packaging options under ProductVariations, each with its own DigiKey
// part number, stock, MOQ and price ladder. V3 flattened all of that onto the
// product, so both shapes are handled.
function variationsOf(product) {
  if (Array.isArray(product.ProductVariations) && product.ProductVariations.length > 0) {
    return product.ProductVariations.map((v) => ({
      supplierPartNumber: v.DigiKeyProductNumber || v.DigiKeyPartNumber || null,
      packaging: (v.PackageType && (v.PackageType.Name || v.PackageType.Value)) || null,
      stock: firstFinite([
        parseQuantity(v.QuantityAvailableforPackageType),
        parseQuantity(v.QuantityAvailable),
      ]),
      minimumOrderQuantity: parseQuantity(v.MinimumOrderQuantity) || 1,
      orderMultiple: parseQuantity(v.StandardPackage) || 1,
      priceBreaks: priceBreaksOf(v.StandardPricing),
      marketPlace: !!v.MarketPlace,
    }));
  }
  return [
    {
      supplierPartNumber: product.DigiKeyPartNumber || product.DigiKeyProductNumber || null,
      packaging: (product.Packaging && (product.Packaging.Name || product.Packaging.Value)) || null,
      stock: parseQuantity(product.QuantityAvailable),
      minimumOrderQuantity: parseQuantity(product.MinimumOrderQuantity) || 1,
      orderMultiple: parseQuantity(product.StandardPackage) || 1,
      priceBreaks: priceBreaksOf(product.StandardPricing),
      marketPlace: false,
    },
  ];
}

function priceBreaksOf(list) {
  if (!Array.isArray(list)) return [];
  return list
    .map((b) => ({
      quantity: parseQuantity(b.BreakQuantity),
      unitPrice: parseMoney(b.UnitPrice),
    }))
    .filter((b) => Number.isFinite(b.quantity) && Number.isFinite(b.unitPrice))
    .sort((a, b) => a.quantity - b.quantity);
}

function firstFinite(values) {
  for (const value of values) {
    if (Number.isFinite(value)) return value;
  }
  return null;
}

function buildRecord(product, part, currency, matchCount) {
  const variations = variationsOf(product);
  const classifications = product.Classifications || {};
  return {
    supplier: SUPPLIER,
    manufacturer: manufacturerOf(product) || null,
    manufacturerPartNumber: mpnOf(product) || null,
    description: descriptionOf(product),
    productUrl: product.ProductUrl || null,
    datasheetUrl: product.DatasheetUrl || null,
    leadTime: product.ManufacturerLeadWeeks || product.LeadStatus || null,
    lifecycle: statusOf(product),
    rohs: classifications.RohsStatus || null,
    totalStock: parseQuantity(product.QuantityAvailable),
    currency,
    matchCount,
    exactMatch: normalizeKey(mpnOf(product)) === normalizeKey(part.mpn),
    variations,
  };
}

module.exports = { DigiKeyClient, SUPPLIER };
