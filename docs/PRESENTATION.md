# Ami readout — run-of-show (5 minutes)

For Bhargava and Shree, presenting before 11 October 2026.

## Files and links

| What | Where |
|---|---|
| Deck, PowerPoint (speaker notes inside) | [Ami-L4-Readout-Bhargava-Shree-2026-10-08.pptx](Ami-L4-Readout-Bhargava-Shree-2026-10-08.pptx) |
| Deck, PDF | [Ami-L4-Readout-Bhargava-Shree-2026-10-08.pdf](Ami-L4-Readout-Bhargava-Shree-2026-10-08.pdf) |
| Deck, live and editable | https://claude.ai/artifact/KMns74rheLukvPD6BzQzoW (share from its Share menu; also exports PPTX/PDF) |
| Pull request to review and merge | https://github.com/barusr9/aurex-ami/pull/2 |
| This repo: plan, logs, evidence, code snapshot | https://github.com/shreenathacc22/ami-l4-project |
| The one-page readout | [../code/READOUT.md](../code/READOUT.md) |

## Run-of-show

Total about 5 minutes 20 seconds. Suggested split: Shree opens and closes, Bhargava takes the middle. Swap as you like; the notes under each slide carry the words.

| # | Slide | Who | Time | The one line |
|---|---|---|---|---|
| 1 | Ami, hardened | Shree | 0:15 | We fixed where it only looked like it worked, and measured every fix |
| 2 | The agent looked like it worked | Shree | 0:40 | Cancel/return never completed; a 502 hung a turn 14 min; nothing was measured |
| 3 | Same frozen suites, before and after | Shree | 0:30 | 20 behavioural + 28 golden cases, judge audited first; same model, same day; we report the range |
| 4 | Better answers, at a lower cost | Bhargava | 0:30 | 0.805 → 0.967 · 16 → 24 of 28 · −38% cost · −23% p95 |
| 5 | Before and after, same frozen suites | Bhargava | 0:40 | 17/20 and 20/20 is the same code twice: that is variance, so we show the range |
| 6 | Six changes, each one measured | Bhargava | 0:50 | Two correctness, one resilience, two cost, one safety |
| 7 | Prompt injection, in the real chat UI | Bhargava | 0:30 | Same message: 90 s of "thinking…" vs a 14-second handoff |
| 8 | Cost fell 38 percent. The target was half | Shree | 0:45 | S1 missed; the 12-word cap broke a guardrail and was reverted; routing can't save money on this proxy |
| 9 | What we would watch, and who gets paged | Shree | 0:20 | Six alerts, thresholds in config, /logs verified |
| 10 | Three things before the 11th | Shree | 0:20 | S5 study, witnessed run, PR review; links on the slide |

If you run long: say slide 9 in one sentence and skip to 10. Never cut slide 8; the honest findings are the point of the readout.

## Likely questions, and the answers

**Why not half the cost, as S1 asked?**
Most of each call is a fixed prompt prefix. The proxy caches it automatically but unpredictably: cache hits swung from 79% to 48% between two runs of the same suite. Tokens fell 39%; dollars fell 38% on the golden set and 11% on the behavioural suite. The answer cache halves the cost of repeated questions, but the frozen suites contain no repeats, by design.

**The same code scored 17/20 and 20/20. Which is right?**
Both. The three cases that flip are refusals whose wording sometimes drops the reason ("already shipped", "30-day window"). The model is non-deterministic at this temperature, so we report the range, not the best run. READOUT §3 says this out loud.

**How do you know the LLM judge is trustworthy?**
We audited it before any golden run: given each case's own reference answer it scores 0.98; given an unrelated reference it scores 0.05. It tells all 28 references apart. One row (`tone-upset-customer`) has a borderline reference, and we flag it.

**Was the 14-minute hang real, or a test artefact?**
Real, and reproducible: the proxy returns a 502 on the injection text every time, and the old code retried with hidden SDK retries stacked on its own. We re-created it today in the real chat UI: still "thinking…" at 90 seconds (we stopped there). The new code hands off in about 15 seconds.

**What did model routing achieve?**
Nothing live: the class proxy serves `gpt-5.6-terra` whatever model is requested. The routing logic is built, tested, and shows −25% in simulation. It ships switched off, and we say so.

**Can the answer cache serve a wrong or private answer?**
It only caches a conversation's first message when it is classified public and answered from the knowledge base alone, keyed by text plus a stamp of the knowledge documents, with a TTL. Account questions and failure replies are never cached. One phrasing of "refunds" is classified private and is never cached, which is the classifier being cautious.

**Why hasn't the user study (S5) run?**
It needs three people and a facilitator; the agent cannot run it. The kit, a reset script and a facilitator sheet are ready, and it is the first open item.

**What happens between now and the 11th?**
The S5 sessions, then the witnessed run of both suites on the merged branch (about 10 minutes, 220k tokens) that replaces the table's numbers and gets a signature, then the PR review and merge.

## Before you present

1. Open the PPTX once in PowerPoint and page through it; the speaker notes are under each slide. If anything looks off, the PDF is the exact rendering of the live deck.
2. Have READOUT.md and PR #2 open in browser tabs for questions.
3. Optional live demo (adds about 30 seconds, some risk): with the app running on port 4000, type the injection message and show the handoff reply arriving in about 15 seconds. Skip it if the network is slow that day.
4. Say the numbers predate the merge with master; the witnessed run on the merged branch is the final measurement.
