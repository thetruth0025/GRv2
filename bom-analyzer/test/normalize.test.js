'use strict';

const test = require('node:test');
const assert = require('node:assert');

const {
  LIFECYCLE,
  normalizeLifecycle,
  worstLifecycle,
  parseLeadTimeDays,
  formatLeadTime,
  parseMoney,
  priceAtQuantity,
  orderQuantity,
  pickVariation,
  recordToOffer,
  buildOffer,
  compareOffers,
  missingOffer,
} = require('../lib/normalize');

test('lifecycle vocabulary from both suppliers maps onto one scale', () => {
  assert.equal(normalizeLifecycle('Active'), LIFECYCLE.ACTIVE);
  assert.equal(normalizeLifecycle('Obsolete'), LIFECYCLE.OBSOLETE);
  assert.equal(normalizeLifecycle('Not For New Designs'), LIFECYCLE.NRND);
  assert.equal(normalizeLifecycle('NRND'), LIFECYCLE.NRND);
  assert.equal(normalizeLifecycle('Last Time Buy'), LIFECYCLE.LAST_TIME_BUY);
  assert.equal(normalizeLifecycle('End of Life'), LIFECYCLE.END_OF_LIFE);
  assert.equal(normalizeLifecycle('Discontinued at Digi-Key'), LIFECYCLE.DISCONTINUED);
  assert.equal(normalizeLifecycle('New at Mouser'), LIFECYCLE.NEW);
  assert.equal(normalizeLifecycle(''), LIFECYCLE.UNKNOWN);
  assert.equal(normalizeLifecycle(null), LIFECYCLE.UNKNOWN);
});

test('the worst status across suppliers wins so a risk is never hidden', () => {
  assert.equal(worstLifecycle([LIFECYCLE.ACTIVE, LIFECYCLE.OBSOLETE]), LIFECYCLE.OBSOLETE);
  assert.equal(worstLifecycle([LIFECYCLE.ACTIVE, LIFECYCLE.NRND]), LIFECYCLE.NRND);
  assert.equal(worstLifecycle([LIFECYCLE.ACTIVE, LIFECYCLE.ACTIVE]), LIFECYCLE.ACTIVE);
  assert.equal(worstLifecycle([]), LIFECYCLE.UNKNOWN);
});

test('lead times parse from every unit the suppliers use', () => {
  assert.equal(parseLeadTimeDays('12 Weeks'), 84);
  assert.equal(parseLeadTimeDays('45 Days'), 45);
  assert.equal(parseLeadTimeDays('3 Months'), 90);
  assert.equal(parseLeadTimeDays('In Stock'), 0);
  assert.equal(parseLeadTimeDays(16), 16);
  assert.equal(parseLeadTimeDays(''), null);
  assert.equal(parseLeadTimeDays(null), null);
  // A bare number in a lead-time field means weeks at both suppliers.
  assert.equal(parseLeadTimeDays('20'), 140);
});

test('lead times format back into buyer-readable text', () => {
  assert.equal(formatLeadTime(0), 'In stock');
  assert.equal(formatLeadTime(84), '12 weeks');
  assert.equal(formatLeadTime(7), '1 week');
  assert.equal(formatLeadTime(45), '45 days');
  assert.equal(formatLeadTime(null), null);
});

test('prices parse out of the strings Mouser returns', () => {
  assert.equal(parseMoney('$1.23'), 1.23);
  assert.equal(parseMoney('1.23'), 1.23);
  assert.equal(parseMoney(0.0087), 0.0087);
  assert.equal(parseMoney('$1,234.56'), 1234.56);
  // European formatting: comma is the decimal separator.
  assert.equal(parseMoney('1.234,56 €'), 1234.56);
  assert.equal(parseMoney(''), null);
});

test('the applicable price break is the highest one at or below the quantity', () => {
  const breaks = [
    { quantity: 1, unitPrice: 1.0 },
    { quantity: 10, unitPrice: 0.8 },
    { quantity: 100, unitPrice: 0.5 },
  ];
  assert.equal(priceAtQuantity(breaks, 1).unitPrice, 1.0);
  assert.equal(priceAtQuantity(breaks, 9).unitPrice, 1.0);
  assert.equal(priceAtQuantity(breaks, 10).unitPrice, 0.8);
  assert.equal(priceAtQuantity(breaks, 250).unitPrice, 0.5);
  assert.equal(priceAtQuantity([], 10), null);
});

test('below the smallest break the buyer still pays that break price', () => {
  const breaks = [{ quantity: 10, unitPrice: 0.8 }];
  assert.equal(priceAtQuantity(breaks, 1).unitPrice, 0.8);
});

test('order quantity respects minimum order and packaging multiples', () => {
  assert.equal(orderQuantity(100, 1, 1), 100);
  assert.equal(orderQuantity(5, 10, 1), 10);
  assert.equal(orderQuantity(105, 1, 50), 150);
  assert.equal(orderQuantity(3, 1, 1), 3);
  assert.equal(orderQuantity(0, 1, 1), 1);
});

test('an offer prices the quantity actually purchased, not the quantity needed', () => {
  const offer = buildOffer({
    supplier: 'Mouser',
    quantity: 105,
    minimumOrderQuantity: 1,
    orderMultiple: 50,
    priceBreaks: [
      { quantity: 1, unitPrice: 0.1 },
      { quantity: 100, unitPrice: 0.05 },
    ],
    stock: 5000,
    leadTime: '8 Weeks',
    lifecycle: 'Active',
  });

  assert.equal(offer.orderQuantity, 150);
  assert.equal(offer.unitPrice, 0.05);
  assert.equal(offer.extendedPrice, 7.5);
  assert.equal(offer.stockSufficient, true);
  assert.equal(offer.leadTimeDays, 56);
  assert.equal(offer.lifecycle, LIFECYCLE.ACTIVE);
});

test('stock below the order quantity is reported as insufficient', () => {
  const offer = buildOffer({
    supplier: 'DigiKey',
    quantity: 500,
    stock: 100,
    priceBreaks: [{ quantity: 1, unitPrice: 1 }],
  });
  assert.equal(offer.stockSufficient, false);
});

test('packaging selection prefers an option that covers the quantity, then price', () => {
  const variations = [
    {
      supplierPartNumber: 'CUT-TAPE',
      stock: 50,
      minimumOrderQuantity: 1,
      orderMultiple: 1,
      priceBreaks: [{ quantity: 1, unitPrice: 0.1 }],
    },
    {
      supplierPartNumber: 'REEL',
      stock: 10000,
      minimumOrderQuantity: 1,
      orderMultiple: 1,
      priceBreaks: [{ quantity: 1, unitPrice: 0.12 }],
    },
  ];
  // 50 fits cut tape, which is also cheaper.
  assert.equal(pickVariation(variations, 50).supplierPartNumber, 'CUT-TAPE');
  // 500 does not, so the reel wins despite the higher unit price.
  assert.equal(pickVariation(variations, 500).supplierPartNumber, 'REEL');
});

test('a low unit price does not win when its minimum forces a much larger buy', () => {
  const variations = [
    {
      supplierPartNumber: 'CUT-TAPE',
      stock: 5000,
      minimumOrderQuantity: 1,
      orderMultiple: 1,
      priceBreaks: [{ quantity: 100, unitPrice: 0.0156 }],
    },
    {
      supplierPartNumber: 'REEL',
      stock: 1000000,
      minimumOrderQuantity: 5000,
      orderMultiple: 5000,
      priceBreaks: [{ quantity: 5000, unitPrice: 0.00518 }],
    },
  ];
  // The reel's unit price is a third of cut tape, but buying 500 pieces costs
  // $7.80 on cut tape against $25.90 for a minimum reel.
  assert.equal(pickVariation(variations, 500).supplierPartNumber, 'CUT-TAPE');
  // Past the cut-tape stock the reel is the only option that can ship.
  assert.equal(pickVariation(variations, 20000).supplierPartNumber, 'REEL');
});

test('marketplace packaging is a last resort even when it is cheapest', () => {
  const variations = [
    {
      supplierPartNumber: 'MARKET',
      stock: 10000,
      minimumOrderQuantity: 1,
      orderMultiple: 1,
      priceBreaks: [{ quantity: 1, unitPrice: 0.05 }],
      marketPlace: true,
    },
    {
      supplierPartNumber: 'STOCK',
      stock: 10000,
      minimumOrderQuantity: 1,
      orderMultiple: 1,
      priceBreaks: [{ quantity: 1, unitPrice: 0.09 }],
      marketPlace: false,
    },
  ];
  assert.equal(pickVariation(variations, 100).supplierPartNumber, 'STOCK');
});

test('a cached catalog record reprices at a different quantity without refetching', () => {
  const record = {
    supplier: 'DigiKey',
    manufacturerPartNumber: 'RC0603FR-0710KL',
    leadTime: '10 Weeks',
    lifecycle: 'Active',
    totalStock: 100000,
    currency: 'USD',
    variations: [
      {
        supplierPartNumber: '311-10.0KHRCT-ND',
        stock: 100000,
        minimumOrderQuantity: 1,
        orderMultiple: 1,
        priceBreaks: [
          { quantity: 1, unitPrice: 0.1 },
          { quantity: 1000, unitPrice: 0.01 },
        ],
      },
    ],
  };

  assert.equal(recordToOffer(record, { quantity: 10 }).unitPrice, 0.1);
  assert.equal(recordToOffer(record, { quantity: 5000 }).unitPrice, 0.01);
  assert.equal(recordToOffer(record, { quantity: 5000 }).extendedPrice, 50);
});

// ── Cross-supplier comparison ───────────────────────────────────────────────

function offer(supplier, overrides) {
  return buildOffer(
    Object.assign(
      {
        supplier,
        quantity: 100,
        stock: 1000,
        leadTime: '10 Weeks',
        lifecycle: 'Active',
        priceBreaks: [{ quantity: 1, unitPrice: 1 }],
      },
      overrides
    )
  );
}

test('the cheaper supplier wins on price and is named', () => {
  const summary = compareOffers(
    [
      offer('DigiKey', { priceBreaks: [{ quantity: 1, unitPrice: 1 }] }),
      offer('Mouser', { priceBreaks: [{ quantity: 1, unitPrice: 0.75 }] }),
    ],
    100
  );
  assert.equal(summary.bestPriceSupplier, 'Mouser');
  assert.equal(summary.bestPrice, 75);
  assert.equal(summary.priceSpread, 25);
});

test('stock on hand beats a shorter quoted factory lead time', () => {
  const summary = compareOffers(
    [
      offer('DigiKey', { stock: 0, leadTime: '2 Weeks' }),
      offer('Mouser', { stock: 5000, leadTime: '30 Weeks' }),
    ],
    100
  );
  assert.deepEqual(summary.inStockSuppliers, ['Mouser']);
  assert.equal(summary.bestLeadTimeDays, 0);
  assert.equal(summary.bestLeadTimeSupplier, 'Mouser');
  assert.equal(summary.recommendedSupplier, 'Mouser');
});

test('with both suppliers in stock no lead-time winner is invented', () => {
  const summary = compareOffers(
    [
      offer('DigiKey', { stock: 5000, leadTime: '30 Weeks', priceBreaks: [{ quantity: 1, unitPrice: 1 }] }),
      offer('Mouser', { stock: 5000, leadTime: '2 Weeks', priceBreaks: [{ quantity: 1, unitPrice: 2 }] }),
    ],
    100
  );
  // Both ship today, so neither factory lead time is the one to act on.
  assert.equal(summary.bestLeadTimeDays, 0);
  assert.equal(summary.bestLeadTimeSupplier, null);
  assert.deepEqual(summary.bestLeadTimeSuppliers.sort(), ['DigiKey', 'Mouser']);
  // The cheaper of the two equally fast suppliers is the recommendation.
  assert.equal(summary.recommendedSupplier, 'DigiKey');
});

test('with nobody holding stock the shortest lead time is recommended', () => {
  const summary = compareOffers(
    [
      offer('DigiKey', { stock: 0, leadTime: '20 Weeks' }),
      offer('Mouser', { stock: 0, leadTime: '6 Weeks' }),
    ],
    100
  );
  assert.equal(summary.recommendedSupplier, 'Mouser');
  assert.equal(summary.bestLeadTimeDays, 42);
});

test('a split order is suggested when neither supplier alone can cover the line', () => {
  const summary = compareOffers(
    [offer('DigiKey', { stock: 60 }), offer('Mouser', { stock: 60 })],
    100
  );
  const texts = summary.flags.map((f) => f.text);
  assert.ok(texts.some((t) => /split the order/.test(t)), texts.join(' | '));
});

test('a genuine shortage is called out rather than treated as a split', () => {
  const summary = compareOffers(
    [offer('DigiKey', { stock: 10 }), offer('Mouser', { stock: 5 })],
    100
  );
  const texts = summary.flags.map((f) => f.text);
  assert.ok(texts.some((t) => /below the required 100/.test(t)), texts.join(' | '));
});

test('an obsolete flag from one supplier marks the whole line', () => {
  const summary = compareOffers(
    [offer('DigiKey', { lifecycle: 'Active' }), offer('Mouser', { lifecycle: 'Obsolete' })],
    100
  );
  assert.equal(summary.lifecycle, LIFECYCLE.OBSOLETE);
  assert.equal(summary.lifecycleSeverity, 'bad');
  assert.ok(summary.flags.some((f) => /find a replacement/.test(f.text)));
});

test('a part no supplier carries is flagged rather than silently priced', () => {
  const summary = compareOffers(
    [missingOffer('DigiKey', 'no match'), missingOffer('Mouser', 'no match')],
    100
  );
  assert.equal(summary.recommendedSupplier, null);
  assert.equal(summary.bestPrice, null);
  assert.ok(summary.flags.some((f) => /Not found at any/.test(f.text)));
});

test('a single-source part is called out', () => {
  const summary = compareOffers([offer('DigiKey'), missingOffer('Mouser', 'no match')], 100);
  assert.ok(summary.flags.some((f) => /Single source/.test(f.text)));
});

test('a long lead time is surfaced even when everything else is fine', () => {
  const summary = compareOffers(
    [offer('DigiKey', { stock: 0, leadTime: '40 Weeks' }), offer('Mouser', { stock: 0, leadTime: '52 Weeks' })],
    100
  );
  assert.ok(summary.flags.some((f) => /Best lead time is/.test(f.text)));
});
