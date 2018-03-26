Select Quote.*, Nvl (Labor.laborcost, 0000000.00) As laborcost, Nvl (Labor.laborhrs, 0000000.00) As laborhrs, Nvl (Mtl.mtlcost, 0000000.00) As mtlcost, Nvl (Mtl.mtlmarkup, 00000000.00) As mtlmarkup From Quote Left Join (Select QuoteEst.ifkey, x,y, Cast (Sum (QuoteEst.Qty * QuoteEst.Time * QuoteEst.mach_rate * Quote.labormult / Iif (QuoteEst.per_job, Quote.Qty, 1)) As N(10, 2)) As laborcost, Cast (Sum (QuoteEst.Qty * QuoteEst.Time / Iif (QuoteEst.per_job, Quote.Qty, 1)) As N(10, 2)) As laborhrs From QuoteEst Left Join Quote On QuoteEst.ifkey = Quote.ipkey Group By 1) Labor On Labor.ifkey = Quote.ipkey Left Join (Select QuoteMtl.ifkey, z,x, Cast (Sum (QuoteMtl.Qty * QuoteMtl.price / Iif (QuoteMtl.per_job, Quote.Qty, 1)) As N(10, 2)) As mtlcost, Cast (Sum (QuoteMtl.Qty * QuoteMtl.price * (QuoteMtl.markup - 1) / Iif (QuoteMtl.per_job, Quote.Qty, 1)) As N(10, 2)) As mtlmarkup From QuoteMtl Left Join Quote On QuoteMtl.ifkey = Quote.ipkey Group By 1) Mtl On Mtl.ifkey = Quote.ipkey Where << This.cWhereClause >> 