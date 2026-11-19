# WPS compatibility experiment: gratitude class meeting

This branch contains the WPS configuration changes and observations from one generated deck. It is an experiment for review; no merge into main is implied.

## Input and baseline

- Baseline: `c69e6263459f3ffc4da6dc1ad816da9fbf63f66c`.
- Topic: 第七次主题班会「知恩于心，感恩于行」.
- Audience: university students.
- Planned speaking and activity time: 900 seconds. This is a timing plan, not a measured live presentation.
- Target player: WPS; version unknown.
- Input: `examples/gratitude-wps.json`.
- Route exercised: the existing `deck_builder.py` fallback generator. The template, image-generation and calligraphic-cover routes were not exercised.

## Observed output and targeted changes

| Observation | Owning implementation change |
| --- | --- |
| All 13 original slides emitted `p:morph` in the base namespace | WPS uses fade/none. PowerPoint Morph uses its extension namespace and a fade fallback. |
| Text lacked explicit East Asian font declarations | Add script-specific declarations while preserving explicit Latin and East Asian font choices; do not claim font embedding. |
| Timeline body text was 13pt; some other body text was 14pt | Raise the affected text in the exercised layouts to 18pt, and ordinary headings to 32pt. |
| The check option only printed inspection output | Run layout and WPS static checks and return a failure code for errors. |
| Body text outside the canvas was skipped as decorative bleed | Report body text outside the canvas; keep intentional decorative bleed possible. |
| Speaker notes from the input were discarded | Store the supplied notes in all 13 slides. |
| Preview used Android-only font paths and ignored the slide background | Add Windows font paths/configuration and use the slide background; the preview remains approximate. |

## Results

The same checker was applied to the actual generated files:

| Output | WPS conservative-profile errors | Body layout errors |
| --- | ---: | ---: |
| Original generator | 13 | 2 |
| WPS configuration, first pass | 0 | 2 |
| Revised sample | 0 | 0 |

The revised sample retained the slide count and editable text content. It added 13 pages of speaker notes. Eleven regression tests passed, covering transition fallback, extension markup/order, font preservation, source-file preservation, external-resource detection, CLI failure status, empty specifications, the template route, notes and layout-failure propagation. The skill entrypoint validator also passed.

These are static/profile checks, not an application compatibility certification. No final-file opening, slideshow, or PDF export was verified in WPS. Native WPS status remains `not_performed`.

## Review questions for the next iteration

The sample is entirely native text and code-generated shapes, with no photographs. Several layouts still rely on cards. The cover is the generator fallback, not the repository's elaborate calligraphy/character cover. The generic VS marker may not suit classroom discussion. Footers and decorative numbering retain small sizes and carry no key content.

These observations should guide review of the actual deck before further design-rule changes. They should not become mandatory requirements for unrelated presentation types.

## Reproduce

Install the existing dependencies (`python-pptx`, `Pillow`, and `PyYAML`), then run:

```bash
python scripts/deck_builder.py examples/gratitude-wps.json -o gratitude-wps.pptx --target wps --check
python scripts/test_wps_compat.py
```

Open the final file in the intended WPS version and record the file hash, OS, version and checked pages before describing it as natively verified.
