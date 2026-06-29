# Refinery Outages — Presentation Study Sheet

*One sheet for all three talks (Gasoline · Diesel/Distillate · Chem-Feed). Numbers pulled
from the model, day-weighted concurrent capacity offline (kbd), per unit. As of the latest
Snowflake pull (actuals reported through ~mid-2026; 2027 is the booked forward plan).*

---

## The 20-second version (lead with this)

- We track **capacity offline per unit** — CDU (crude), FCC (cat cracker), hydrocracker, reformer —
  by month, region and operator, **never summed** (a 250-kbd CDU plus a 100-kbd FCC is **not** "350 offline").
- **2027 is heavier than 2026 on a like-for-like (H1 planned) basis across all four units**, landing into a
  **thin-inventory backdrop** (US stocks vs 5-yr avg: **gasoline −5%, distillate −10%, crude −7%**).
- It is **not a record** — the all-in figure looks lighter than 2023–25 only because 2027 has **no unplanned
  actuals yet** and **non-Exxon H2 isn't fully booked**. It's still filling in.
- The trade-relevant squeeze: **spring crude turnarounds** cut the whole barrel right as the **summer-grade
  gasoline switchover** starts (Mar–Jun), with **distillate inventories the tightest of the three**.

---

## Universal facts (true in every talk)

- **The four units & what they make:** CDU = crude (the whole barrel — gasoline *and* distillate);
  FCC = cat gasoline + octane; **hydrocracker = diesel & jet**; reformer = reformate (the octane in gasoline,
  and the aromatics/petchem feed).
- **Per unit, never summed; day-weighted.** A unit down part of a month counts only for its days down
  (nameplate × days-down ÷ days-in-month); each unit once per month.
- **2027 completeness is asymmetric.** Only **ExxonMobil gave a full-year plan** (verified against their
  corporate schedule). **Every other operator is H1-confirmed only**, so non-Exxon H2 (the autumn bars) is an
  **indicative floor, not booked** — don't trade it as firm.
- **CDU = atmospheric crude only.** Vacuum (VDU) is excluded, even where the source mislabels it.
- **Source:** Snowflake golden record (2023 → 2027), one pipeline behind the decks, the Excel model and the
  dashboard — so all three agree.

### Quick-reference numbers (2027, kbd)

| Unit | 2027 confirmed peak month | H1 planned vs 2026 | vs 2025 | Where the work sits (by PADD) |
|---|---|---|---|---|
| **CDU** (crude) | ~1,060 (Mar)* | **+7%** | below 2025 | **P3 Gulf** ≫ P2 Midwest > P5 West |
| **FCC** (gasoline) | ~450 (Feb) | **+66%** | above 2025 | **P3 Gulf** ≫ P2 Midwest |
| **Hydrocracker** (diesel) | ~140 (Feb) | **+9%** | above 2025 | **P3 Gulf** > P5 West |
| **Reformer** (octane) | ~185 (Mar) | **+38%** | below 2025 | **P2 Midwest** + P3 Gulf |

\* *Confirmed (H1) peak is March; the **modeled full-year peak is October**, but October leans on
non-Exxon H2 that isn't booked yet. If asked "when's the peak," say: confirmed = spring; the fall stack is
indicative.*

**Biggest single 2027 outages (all crude/CDU):** Joliet ~250 (P2), Cherry Point ~250 (P5), Norco ~240 (P3),
Richmond ~240 (P5), Lemont ~175 (P2), Martinez ~155 (P5), Garyville ~145 (P3). By tonnage the biggest single
outages lean **P5 (West) and P3 (Gulf)**; total offline is **P3-dominated**.

---

## GASOLINE focus  (main deck — slides 2, 5, 6, 9, 10)

**Units: CDU + FCC.** The story is the gasoline pool tightening into summer.

- **FCC is the standout:** H1 2027 cat-cracker offline is **+66% vs 2026** and above 2025 — the most elevated
  of the four units. P3 (Gulf) carries the bulk.
- **Crude (CDU) +7% H1 vs 2026**, concentrated in **PADD 3 (Gulf)**, the export/USGC swing region.
- **Timing is the point:** crude + cat turnarounds cluster **Feb–May**, landing into the **summer-grade
  switchover (Mar–Jun)** — supply tightens just as summer-grade demand builds.
- **Backdrop:** gasoline stocks **−5% vs the 5-yr average**.
- **Octane angle (hand to the chem-feed read):** reformer turnarounds cut reformate, squeezing the gasoline
  pool's octane even when crude runs hold.

**Likely questions**
- *"Is gasoline supply actually tight or just turnarounds?"* — Both: −5% inventory cushion **and** the heaviest
  FCC turnaround H1 in three years, into the summer-grade flip.
- *"Which region matters?"* — PADD 3 (Gulf): it's the swing for crude + cat and drives the export barrel.
  PADD 5 (West) is islanded — a California outage isn't bailed out.
- *"Biggest single risk?"* — A Gulf CDU trip in the spring window; crude down cuts cat-gasoline feed too.

---

## DIESEL / DISTILLATE focus  (main deck — slides 2, 5, 7, 9, 10)

**Units: CDU + hydrocracker.** The hydrocracker is small in kbd, so the diesel story is **crude + inventories**,
not hydrocracker turnarounds alone.

- **Distillate inventories are the tightest of the three: −10% vs the 5-yr average.** This is the diesel
  headline — lead with it.
- **Crude (CDU) cuts distillate too.** Every CDU outage removes the whole barrel; the same spring crude
  turnarounds that hit gasoline also pull distillate yield. CDU is **P3 Gulf-heavy**.
- **Hydrocracker itself is modest** (~140 kbd peak, H1 +9% vs 2026 — gently rising, above 2025 but not
  elevated), front-loaded **Q1** plus a **Sep–Oct** window; **P3 Gulf** then **P5 West**.
- **Risk window is winter** (heating/diesel demand) on top of a thin cushion.

**Likely questions**
- *"Hydrocracker outages look small — why care?"* — The diesel squeeze is **inventory + crude**, not the
  hydrocracker line. Distillate stocks are −10% and crude turnarounds cut distillate yield directly.
- *"Is 2027 diesel risk worse than 2026?"* — H1 hydrocracker is **+9% vs 2026** and crude is **+7%**, into a
  tighter distillate cushion — so directionally yes, but it's the inventory backdrop doing the work.
- *"Where's the diesel exposure?"* — PADD 3 (Gulf) for both crude and hydrocracker; watch the Q1 and fall windows.

---

## CHEM-FEED focus  (naphtha deck — all 8 slides; window = rest of 2026 + 2027)

**Units: reformer + the naphtha complex.** Octane and petrochemical feedstock, not crude yields.

- **Window is the forward book: rest of 2026 + all of 2027** (the senior analyst's ask), 2026 H1 shaded as actual.
- **Naphtha runs structurally short:** 2027 net **−1,347 kbd, 11 of 12 months in deficit** (tightest **Oct −348,
  Sep −230**); 2026 net **−1,119 kbd, 10 deficit months** (tightest **Apr −340, Feb −226**). Crude turnarounds
  pull more naphtha *off* than reformer turnarounds free up → **bullish reformate / octane**.
- **Reformer / octane:** across the forward window ~**1,230 kbd-months** of reformer offline ≈ **~1,040 kbd-months
  of reformate (octane) not made**. P3 (Gulf) carries the most; **P2 (Midwest)** is the other pole.
- **The mechanic both ways:** a CDU outage removes naphtha *supply*; a reformer outage removes *demand* and cuts
  ~85% of its charge as reformate — squeezing the gasoline pool's octane and the aromatics/BTX petchem feed.

**Likely questions**
- *"Long or short naphtha?"* — **Short**, nearly every month, both years. Crude down > reformer down.
- *"Why does a CDU-only tracker miss this?"* — It sees crude runs but not that octane and chem-feed tighten when
  reformers/naphtha move; this deck is the octane + petrochemical-feed read.
- *"Tightest months?"* — 2026: **Apr / Feb**; 2027: **Oct / Sep** (the autumn crude stack).

---

## Month-over-month — the "what's changed" read  (main deck slide 4)

- Latest reported month vs prior: **Jul 2026 vs Jun 2026**, total offline **1,075 vs 1,350 kbd (−275)** —
  falling into summer (driving-season, turnarounds wind down).
- By region: **PADD 1 +320** (more offline), **PADD 3 −477**, PADD 2 −100.
- By unit: **FCC −77, CDU −56, reformer −32, hydrocracker +8**.
- **5 outages newly appeared, 48 resolved / came back** month-on-month.
- *Heads-up:* "July" shows as the latest month because the pull includes early-July reporting — it's the latest
  month in the book, not a future guess. (Balances are read month-over-month because contracts price monthly.)

---

## Anticipated questions — crisp answers (cross-cutting)

- **"Is 2027 a record year for outages?"** — No. 2027 H1 planned is up vs 2026 (FCC +66%, reformer +38%, CDU +7%,
  HCU +9%), and FCC/hydrocracker are above 2025 too — but the **all-in** figure looks lighter than 2023–25 only
  because 2027 has **no unplanned actuals yet** and **non-Exxon H2 isn't booked**. Heavier H1, still filling in.
- **"Why per-unit, never a single total?"** — Different units make different products; adding a CDU to an FCC is
  meaningless. We read each on its own.
- **"How firm is the 2027 plan?"** — H1 is the honest cross-year window. Only ExxonMobil is full-year-verified;
  everyone else is H1-confirmed, so non-Exxon H2 is an indicative floor that grows as operators book.
- **"What's the unplanned scenario?"** — A monthly **risk range, not a forecast**: the mean 2023–26 monthly
  shape scaled ×0.8 / ×1.0 / ×1.3. Average path peaks ~**2,880 kbd** in the worst month, Active ~**3,740**; risk
  peaks in **February** (freeze) and **Sep–Oct** (turnaround overlap). Don't read it as an annual total.
- **"Where's the data from / how current?"** — Snowflake golden record, actuals through ~mid-2026, 2027 = booked
  plan. Same source behind the deck, the Excel model and the dashboard.
- **"Can you back any number out?"** — Yes; every figure is a visible `=SUMIFS` in the Excel model (12 tabs),
  traceable to the source `Data` sheet.

---

## Landmines (don't get caught)

- **Don't sum across units** and don't quote an "annual total" of offline — it's a sum of 12 monthly concurrent
  figures, a magnitude, not a level. Talk **peak month** and **average month**.
- **Don't present non-Exxon H2 as confirmed** — the autumn bars past the dotted line are indicative.
- **October ≠ confirmed peak.** The modeled CDU peak is October, but it leans on unbooked H2; the confirmed peak
  is **March**.
- **If a refinery appears twice** (e.g., Garyville): it may be **two separate crude trains** *or* an overlap —
  check the model's **Data Quality** tab, which flags the handful of double-counts (~180 kbd across 2026–27,
  minor vs the totals). Don't assert one big number without checking.
- **"Heavier H1" is unit-specific:** FCC & hydrocracker are above both 2025 and 2026; **CDU and reformer are
  above 2026 but below 2025**. Be precise if pressed.
- **MoM anchors on the latest reported month (Jul 2026)** — not a forecast; it's the newest month in the pull.
