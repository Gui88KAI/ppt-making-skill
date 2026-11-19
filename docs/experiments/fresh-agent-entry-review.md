# Fresh-agent entry and image-route clarification

This experiment follows the user's observation that the previous output skipped existing type questions and image generation. The original skill already required those steps. That execution failure should not be described as evidence that the skill cannot generate images.

## Narrow changes

- Make the skill description cover actual authoring requests and exclude explanation/code-review-only requests.
- Put one intake authority at the top: type, topic, audience, occasion, duration and total-page policy. Ask only for missing decisions; never silently default to a competition deck.
- Distinguish fixed pages, maximum pages, a page range and explicit delegation. Suggested type ranges are not user authorization. Covers, endings and opening frames count toward the total.
- Route to one selected type file after intake; move the long retained production handbook into a supporting reference so the initial questionnaire is not buried.
- Add repository AGENTS routing and Codex's implicit-invocation policy without claiming universal automatic loading in every AI host.
- Prefer an available native image-generation tool. Lack of a script API key does not establish that native image generation is unavailable. Save and bind actual output images; do not count a prompt list or text-card draft as a completed visual workflow.
- Separate theme-class-meeting structure from coursework/research-result structure. User visual references outrank generic cover presets; reference duration does not replace the current brief.

## Evidence

The supplied comparison deck contains 15 slides and 14 picture objects: ten topic illustrations and four chapter-preview thumbnails. Some quote/action slides deliberately have no image. Its repeated letter motif and coordinated illustration style matter more than an arbitrary image or shape quota. No raw user deck or extracted artwork is added to this branch.

A native built-in image-generation call was exercised for a new watercolor university-student letter-writing illustration. One image was generated and saved locally. A one-page pilot embeds that exact image by hash and retains editable text. This is one visual probe, not a complete deck or completed three-candidate selection cycle. The original proposed university audience and approximately 15-minute duration remain distinct from the reference deck's 40-minute wording.

## Validation

Ten intake routing tests passed: incomplete topic-only prompts, reuse of known facts, complete briefs, all five type files, duration versus page delegation, explicit counts outside suggestions, maximum/range/delegation policies, invalid limits, invalid types and input preservation. The existing eleven WPS regression checks and the skill entrypoint validator also passed.

These are deterministic routing and file checks, not a claim that an independently tested unfamiliar model always follows the skill. Automatic discovery requires installation and host support. Native WPS slideshow validation remains unperformed.
