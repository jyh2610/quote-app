// Pure calculation helpers shared between the quote editor (App.jsx) and the
// admin stats page (AdminPage.jsx), so the two can never disagree on totals.

export const num = (v) => {
  const n = parseFloat(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
};

export const won = (v) =>
  Math.round(v).toLocaleString("ko-KR", { maximumFractionDigits: 0 });

export function itemAmount(item, orderQty) {
  const amt = num(item.price) * num(item.qty);
  return item.amortize ? amt / (num(orderQty) || 1) : amt;
}

// productionCost/marginAmount/supplyAmount for one quote's saved state.
export function computeQuoteTotals(header, groups, freight, marginRate) {
  const groupSubtotal = (g) =>
    (g.items || []).reduce((sum, it) => sum + itemAmount(it, header?.orderQty), 0);

  const productionCost = (groups || []).reduce((sum, g) => sum + groupSubtotal(g), 0);
  const freightNum = num(freight);
  const marginAmount = (productionCost + freightNum) * (num(marginRate) / 100);
  const supplyAmount = productionCost + freightNum + marginAmount;

  return { productionCost, freightNum, marginAmount, supplyAmount };
}
