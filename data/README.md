# Briefly data

The site (`index.html`) holds no news. It reads **`briefly.xlsx`** here when it opens, then checks **`version.json`** every 5 minutes and offers readers the new content when the version changes. GitHub Pages caches files for about 10 minutes, so a push usually shows up on the site within 10–15 minutes.

## Updating

Never edit cells by hand from a script. Use the tool, which validates everything, keeps at most 7 editions, proves the spreadsheet reproduces the JSON exactly, and rewrites `version.json`:

```
python3 tools/briefly_data.py export data/briefly.xlsx editions.json   # spreadsheet -> JSON {"editions":[newest first]}
# ...edit editions.json...
python3 tools/briefly_data.py import editions.json data/briefly.xlsx   # JSON -> spreadsheet + version.json
git add data/briefly.xlsx data/version.json && git commit -m "Briefly edition <label>" && git push origin main
```

## Schema (version 1)

Every sheet is one Excel table with a fixed header row. Rows belong to an edition through `edition_date` (YYYY-MM-DD). `seq` sets the order. Short lists inside one cell are separated by ` | `.

| Sheet | One row is | Columns |
|---|---|---|
| `editions` | an edition | date, no, label, greeting, para_1–3, hive_letters, hive_centre, hive_words (list), timeline_prompt, debate_motion, debate_for_1–2, debate_against_1–2, debate_opinion |
| `stories` | a brief story, curated read or topic-desk item (`section` = brief / curated / desk) | edition_date, section, seq, id, theme, head, summary, why, cls, stat_big, stat_label, minutes, dek, body_1–4, note, checks (list), think (list), game_type, game_word, game_hint |
| `sources` | a source link for a story | edition_date, item_id, seq, outlet, url |
| `pairs` | a key word of a brief story, or a pair in a Match game | edition_date, parent_id, seq, term, definition |
| `opinions` | an opinion column | edition_date, seq, id, theme, outlet, author, role, date, head, url, gist, other, think (list) |
| `games` | a Briefle, Scramble or Match game | edition_date, seq, id, type, word, hint |
| `connections` | a Connections group | edition_date, seq, name, words (list of 4) |
| `timeline` | a Timeline event | edition_date, seq, event, year |
| `duel` | a Number Duel round | edition_date, seq, a_label, a_value, b_label, b_value, note |
| `audio` | one spoken line (`track` = brief, curated, opinion, games, topic:THEME) | edition_date, track, seq, speaker (F/M), text, tone |
| `meta` | a setting | key, value |

Themes: geopolitics, indiapol, environment, science, tech, economy, literature, history, arts, sport.
