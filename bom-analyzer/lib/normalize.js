'use strict';

// Every supplier reports lead time, lifecycle and pricing in its own shape and
// vocabulary. Everything in this module exists to flatten those into one
// comparable record so the UI can put DigiKey and Mouser side by side.

const LIFECYCLE = {
  ACTIVE: 'Active',
  NRND: 'Not Recommended for New Designs',
  LAST_TIME_BUY: 'Last Time Buy',
  END_OF_LIFE: 'End of Life',
  OBSOLETE: 'Obsolete',
  DISCONTINUED: 'Discontinued',
  PREVIEW: 'Preview',
  NEW: 'New Product',
  UNKNOWN: 'Unknown',
};

// Lower rank is healthier. Used to pick the worst status across suppliers so a
// part flagged obsolete by either one is never shown as simply "Active".
const LIFECYCLE_RANK = {
  [LIFECYCLE.ACTIVE]: 0,
  [LIFECYCLE.NEW]: 0,
  [LIFECYCLE.PREVIEW]: 1,
  [LIFECYCLE.UNKNOWN]: 2,
  [LIFECYCLE.NRND]: 3,
  [LIFECYCLE.LAST_TIME_BUY]: 4,
  [LIFECYCLE.END_OF_LIFE]: 5,
  [LIFECYCLE.DISCONTINUED]: 6,
  [LIFECYCLE.OBSOLETE]: 7,
};

const LIFECYCLE_SEVERITY = {
  [LIFECYCLE.ACTIVE]: 'ok',
  [LIFECYCLE.NEW]: 'ok',
  [LIFECYCLE.PREVIEW]: 'info',
  [LIFECYCLE.UNKNOWN]: 'unknown',
  [LIFECYCLE.NRND]: 'warn',
  [LIFECYCLE.LAST_TIME_BUY]: 'warn',
  [LIFECYCLE.END_OF_LIFE]: 'bad',
  [LIFECYCLE.DISCONTINUED]: 'bad',
  [LIFECYCLE.OBSOLETE]: 'bad',
};

function normalizeLifecycle(raw) {
  if (raw === null || raw === undefined) return LIFECYCLE.UNKNOWN;
  const text = String(raw).trim().toLowerCase();
  if (!text) return LIFECYCLE.UNKNOWN;

  if (/obsolete/.test(text)) return LIFECYCLE.OBSOLETE;
  if (/last\s*time\s*buy|\bltb\b/.test(text)) return LIFECYCLE.LAST_TIME_BUY;
  if (/end\s*of\s*life|\beol\b/.test(text)) return LIFECYCLE.END_OF_LIFE;
  if (/discontinu/.test(text)) return LIFECYCLE.DISCONTINUED;
  // DigiKey says "Not For New Designs", Mouser "Not Recommended for New Designs".
  if (/not\s*(recommended|for\s*new)|\bnrnd\b|no\s*longer\s*manufactured/.test(text)) {
    return LIFECYCLE.NRND;
  }
  if (/preliminary|preview|pre-?release/.test(text)) return LIFECYCLE.PREVIEW;
  if (/new\s*(product|at)/.test(text)) return LIFECYCLE.NEW;
  if (/active|production|normally\s*stocking/.test(text)) return LIFECYCLE.ACTIVE;
  return LIFECYCLE.UNKNOWN;
}

function worstLifecycle(statuses) {
  let worst = null;
  for (const status of statuses) {
    if (!status) continue;
    const rank = LIFECYCLE_RANK[status];
    if (rank === undefined) continue;
    if (worst === null || rank > LIFECYCLE_RANK[worst]) worst = status;
  }
  return worst || LIFECYCLE.UNKNOWN;
}

function lifecycleSeverity(status) {
  return LIFECYCLE_SEVERITY[status] || 'unknown';
}

// Suppliers express lead time as free text: "12 Weeks", "8 weeks", "45 Days",
// "In Stock", "3 mo". Days is the only unit that compares cleanly.
function parseLeadTimeDays(raw) {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'number' && Number.isFinite(raw)) {
    return raw > 0 ? Math.round(raw) : null;
  }
  const text = String(raw).trim().toLowerCase();
  if (!text) return null;
  if (/^(in\s*stock|stock|immediate|available)$/.test(text)) return 0;

  const match = text.match(/(\d+(?:\.\d+)?)\s*([a-z]*)/);
  if (!match) return null;
  const value = parseFloat(match[1]);
  if (!Number.isFinite(value) || value <= 0) return null;
  const unit = match[2];

  if (/^w/.test(unit)) return Math.round(value * 7);
  if (/^m/.test(unit)) return Math.round(value * 30);
  if (/^y/.test(unit)) return Math.round(value * 365);
  if (/^d/.test(unit)) return Math.round(value);
  // A bare number in a lead-time field is conventionally weeks at both
  // suppliers, which is also the safer assumption to surface to a buyer.
  return Math.round(value * 7);
}

function formatLeadTime(days) {
  if (days === null || days === undefined) return null;
  if (days === 0) return 'In stock';
  if (days % 7 === 0 && days >= 7) {
    const weeks = days / 7;
    return weeks === 1 ? '1 week' : weeks + ' weeks';
  }
  return days === 1 ? '1 day' : days + ' days';
}

function parseMoney(raw) {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'number') return Number.isFinite(raw) ? raw : null;
  // Mouser returns strings like "$1.23" or "1,23 €" depending on locale.
  let text = String(raw).trim();
  if (!text) return null;
  text = text.replace(/[^0-9.,-]/g, '');
  if (!text) return null;
  const lastComma = text.lastIndexOf(',');
  const lastDot = text.lastIndexOf('.');
  if (lastComma > lastDot) {
    text = text.replace(/\./g, '').replace(',', '.');
  } else {
    text = text.replace(/,/g, '');
  }
  const value = parseFloat(text);
  return Number.isFinite(value) ? value : null;
}

function parseQuantity(raw) {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'number') return Number.isFinite(raw) ? Math.trunc(raw) : null;
  const text = String(raw).replace(/[^0-9-]/g, '');
  if (!text) return null;
  const value = parseInt(text, 10);
  return Number.isFinite(value) ? value : null;
}

// Price breaks are quantity tiers: the applicable one is the highest break at
// or below the quantity being bought.
function priceAtQuantity(priceBreaks, quantity) {
  if (!Array.isArray(priceBreaks) || priceBreaks.length === 0) return null;
  const qty = Number.isFinite(quantity) && quantity > 0 ? quantity : 1;
  const sorted = priceBreaks
    .filter((b) => b && Number.isFinite(b.quantity) && Number.isFinite(b.unitPrice))
    .sort((a, b) => a.quantity - b.quantity);
  if (sorted.length === 0) return null;

  let chosen = null;
  for (const brk of sorted) {
    if (brk.quantity <= qty) chosen = brk;
  }
  // Below the smallest break the buyer still pays that break's unit price.
  if (!chosen) chosen = sorted[0];
  return chosen;
}

// Suppliers sell in packaging multiples, so the quantity actually purchased can
// exceed what the BOM calls for.
function orderQuantity(needed, moq, multiple) {
  const qty = Number.isFinite(needed) && needed > 0 ? needed : 1;
  const min = Number.isFinite(moq) && moq > 0 ? moq : 1;
  const mult = Number.isFinite(multiple) && multiple > 0 ? multiple : 1;
  let order = Math.max(qty, min);
  if (mult > 1) order = Math.ceil(order / mult) * mult;
  return order;
}

// One supplier's answer for one BOM line, in the shape the frontend renders.
function buildOffer(input) {
  const quantity = Number.isFinite(input.quantity) && input.quantity > 0 ? input.quantity : 1;
  const priceBreaks = Array.isArray(input.priceBreaks) ? input.priceBreaks : [];
  const moq = Number.isFinite(input.minimumOrderQuantity) ? input.minimumOrderQuantity : 1;
  const multiple = Number.isFinite(input.orderMultiple) ? input.orderMultiple : 1;
  const orderQty = orderQuantity(quantity, moq, multiple);
  const brk = priceAtQuantity(priceBreaks, orderQty);
  const unitPrice = brk ? brk.unitPrice : null;
  const extendedPrice = unitPrice === null ? null : round(unitPrice * orderQty, 4);
  const leadTimeDays = parseLeadTimeDays(input.leadTime);
  const stock = Number.isFinite(input.stock) ? input.stock : null;
  const lifecycle = normalizeLifecycle(input.lifecycle);

  return {
    supplier: input.supplier,
    found: true,
    supplierPartNumber: input.supplierPartNumber || null,
    manufacturer: input.manufacturer || null,
    manufacturerPartNumber: input.manufacturerPartNumber || null,
    description: input.description || null,
    productUrl: input.productUrl || null,
    datasheetUrl: input.datasheetUrl || null,
    packaging: input.packaging || null,
    stock,
    stockSufficient: stock === null ? null : stock >= orderQty,
    leadTimeDays,
    leadTimeText: formatLeadTime(leadTimeDays) || (input.leadTime ? String(input.leadTime) : null),
    leadTimeRaw: input.leadTime === undefined ? null : input.leadTime,
    lifecycle,
    lifecycleRaw: input.lifecycle === undefined ? null : input.lifecycle,
    lifecycleSeverity: lifecycleSeverity(lifecycle),
    rohs: input.rohs || null,
    minimumOrderQuantity: moq,
    orderMultiple: multiple,
    orderQuantity: orderQty,
    unitPrice,
    priceBreakQuantity: brk ? brk.quantity : null,
    extendedPrice,
    currency: input.currency || 'USD',
    priceBreaks,
    // How many alternates the supplier returned, so the UI can say when the
    // match was picked out of several candidates.
    matchCount: Number.isFinite(input.matchCount) ? input.matchCount : 1,
    exactMatch: input.exactMatch !== false,
  };
}

// Suppliers sell the same part in several packaging options (cut tape, reel,
// tube), each with its own stock, MOQ and price ladder. Pick the one that can
// actually ship the quantity needed, then the one that costs least to buy.
//
// Ranking has to be on the cost of the whole order, not on unit price: a reel
// often has the lower unit price but a 5,000-piece minimum, so it is the more
// expensive way to obtain the 500 the BOM asked for. Marketplace listings are
// a last resort because they ship separately from the main order.
function pickVariation(variations, needed) {
  const list = Array.isArray(variations) ? variations.filter(Boolean) : [];
  if (list.length === 0) return null;
  const qty = Number.isFinite(needed) && needed > 0 ? needed : 1;

  const scored = list.map((variation) => {
    const orderQty = orderQuantity(qty, variation.minimumOrderQuantity, variation.orderMultiple);
    const brk = priceAtQuantity(variation.priceBreaks, orderQty);
    return {
      variation,
      orderCost: brk ? brk.unitPrice * orderQty : null,
      covers: Number.isFinite(variation.stock) && variation.stock >= orderQty,
    };
  });

  const priced = scored.filter((s) => Number.isFinite(s.orderCost));
  const pool = priced.length > 0 ? priced : scored;
  pool.sort((a, b) => {
    if (a.covers !== b.covers) return a.covers ? -1 : 1;
    if (!!a.variation.marketPlace !== !!b.variation.marketPlace) return a.variation.marketPlace ? 1 : -1;
    const ac = Number.isFinite(a.orderCost) ? a.orderCost : Infinity;
    const bc = Number.isFinite(b.orderCost) ? b.orderCost : Infinity;
    if (ac !== bc) return ac - bc;
    return (b.variation.stock || 0) - (a.variation.stock || 0);
  });
  return pool[0].variation;
}

// A catalog record is the supplier-agnostic, quantity-independent form of one
// product. Caching that instead of a finished offer means a re-run at a
// different quantity reprices from cache without another API call.
function recordToOffer(record, part) {
  if (!record) return null;
  const variation = pickVariation(record.variations, part.quantity);
  const offer = buildOffer({
    supplier: record.supplier,
    supplierPartNumber: variation ? variation.supplierPartNumber : record.supplierPartNumber,
    manufacturer: record.manufacturer,
    manufacturerPartNumber: record.manufacturerPartNumber,
    description: record.description,
    productUrl: record.productUrl,
    datasheetUrl: record.datasheetUrl,
    packaging: variation ? variation.packaging : null,
    stock: variation && Number.isFinite(variation.stock) ? variation.stock : record.totalStock,
    leadTime: record.leadTime,
    lifecycle: record.lifecycle,
    rohs: record.rohs,
    quantity: part.quantity,
    minimumOrderQuantity: variation ? variation.minimumOrderQuantity : 1,
    orderMultiple: variation ? variation.orderMultiple : 1,
    priceBreaks: variation ? variation.priceBreaks : [],
    currency: record.currency,
    matchCount: record.matchCount,
    exactMatch: record.exactMatch,
  });

  offer.totalStock = Number.isFinite(record.totalStock) ? record.totalStock : offer.stock;
  offer.packagingOptions = Array.isArray(record.variations) ? record.variations.length : 1;
  if (Number.isFinite(record.factoryStock)) offer.factoryStock = record.factoryStock;
  if (record.suggestedReplacement) offer.suggestedReplacement = record.suggestedReplacement;
  return offer;
}

function missingOffer(supplier, reason) {
  return {
    supplier,
    found: false,
    reason: reason || 'No match found',
    stock: null,
    leadTimeDays: null,
    leadTimeText: null,
    lifecycle: LIFECYCLE.UNKNOWN,
    lifecycleSeverity: 'unknown',
    unitPrice: null,
    extendedPrice: null,
    priceBreaks: [],
  };
}

function errorOffer(supplier, message) {
  const offer = missingOffer(supplier, message);
  offer.error = true;
  return offer;
}

function round(value, decimals) {
  const factor = Math.pow(10, decimals);
  return Math.round(value * factor) / factor;
}

// Cross-supplier verdict for one BOM line: which supplier wins on price, on
// lead time, on stock, and what the line's overall risk is.
function compareOffers(offers, quantity) {
  const usable = offers.filter((o) => o && o.found);
  const summary = {
    bestPriceSupplier: null,
    bestPrice: null,
    priceSpread: null,
    bestLeadTimeSupplier: null,
    bestLeadTimeDays: null,
    inStockSuppliers: [],
    lifecycle: worstLifecycle(usable.map((o) => o.lifecycle)),
    recommendedSupplier: null,
    flags: [],
  };
  summary.lifecycleSeverity = lifecycleSeverity(summary.lifecycle);

  if (usable.length === 0) {
    summary.flags.push({ level: 'bad', text: 'Not found at any configured supplier' });
    return summary;
  }

  const priced = usable.filter((o) => Number.isFinite(o.extendedPrice));
  if (priced.length > 0) {
    priced.sort((a, b) => a.extendedPrice - b.extendedPrice);
    summary.bestPriceSupplier = priced[0].supplier;
    summary.bestPrice = priced[0].extendedPrice;
    if (priced.length > 1) {
      const worstPrice = priced[priced.length - 1].extendedPrice;
      summary.priceSpread = round(worstPrice - priced[0].extendedPrice, 4);
      summary.priceSpreadPercent =
        priced[0].extendedPrice > 0
          ? round(((worstPrice - priced[0].extendedPrice) / priced[0].extendedPrice) * 100, 1)
          : null;
    }
  }

  const stocked = usable.filter((o) => o.stockSufficient === true);
  summary.inStockSuppliers = stocked.map((o) => o.supplier);

  // How soon the parts can actually arrive: a supplier holding enough stock
  // ships now regardless of what the factory quotes behind it.
  const effective = usable
    .map((offer) => ({
      offer,
      days: offer.stockSufficient === true
        ? 0
        : Number.isFinite(offer.leadTimeDays) ? offer.leadTimeDays : null,
    }))
    .filter((entry) => entry.days !== null);

  if (effective.length > 0) {
    const fastest = Math.min.apply(null, effective.map((entry) => entry.days));
    const winners = effective.filter((entry) => entry.days === fastest);
    summary.bestLeadTimeDays = fastest;
    summary.bestLeadTimeSuppliers = winners.map((entry) => entry.offer.supplier);
    // With every supplier equally fast there is nothing to single out, so the
    // UI gets no winner to badge rather than an arbitrary one.
    summary.bestLeadTimeSupplier =
      winners.length < usable.length && winners.length === 1 ? winners[0].offer.supplier : null;

    const cheapestFastest = winners
      .filter((entry) => Number.isFinite(entry.offer.extendedPrice))
      .sort((a, b) => a.offer.extendedPrice - b.offer.extendedPrice)[0];
    summary.recommendedSupplier = (cheapestFastest || winners[0]).offer.supplier;
  } else {
    summary.bestLeadTimeSuppliers = [];
    summary.recommendedSupplier = summary.bestPriceSupplier || usable[0].supplier;
  }

  const needed = Number.isFinite(quantity) && quantity > 0 ? quantity : 1;
  const totalStock = usable.reduce((sum, o) => sum + (Number.isFinite(o.stock) ? o.stock : 0), 0);

  if (stocked.length === 0) {
    if (totalStock >= needed) {
      summary.flags.push({
        level: 'warn',
        text: 'No single supplier can cover ' + needed + ' — split the order',
      });
    } else {
      summary.flags.push({
        level: 'bad',
        text: 'Combined stock (' + totalStock + ') is below the required ' + needed,
      });
    }
  }
  if (summary.lifecycleSeverity === 'bad') {
    summary.flags.push({ level: 'bad', text: summary.lifecycle + ' — find a replacement' });
  } else if (summary.lifecycleSeverity === 'warn') {
    summary.flags.push({ level: 'warn', text: summary.lifecycle });
  }
  if (Number.isFinite(summary.bestLeadTimeDays) && summary.bestLeadTimeDays >= 84) {
    summary.flags.push({
      level: 'warn',
      text: 'Best lead time is ' + formatLeadTime(summary.bestLeadTimeDays),
    });
  }
  if (usable.length === 1 && offers.length > 1) {
    summary.flags.push({ level: 'info', text: 'Single source — only ' + usable[0].supplier + ' carries it' });
  }
  if (Number.isFinite(summary.priceSpreadPercent) && summary.priceSpreadPercent >= 15) {
    summary.flags.push({
      level: 'info',
      text: summary.priceSpreadPercent + '% price spread between suppliers',
    });
  }

  return summary;
}

module.exports = {
  LIFECYCLE,
  normalizeLifecycle,
  worstLifecycle,
  lifecycleSeverity,
  parseLeadTimeDays,
  formatLeadTime,
  parseMoney,
  parseQuantity,
  priceAtQuantity,
  orderQuantity,
  pickVariation,
  recordToOffer,
  buildOffer,
  missingOffer,
  errorOffer,
  compareOffers,
  round,
};
