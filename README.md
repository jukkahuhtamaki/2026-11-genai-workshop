# Case template: responsible use of GenAI in research

Slides for a workshop on the responsible use of generative AI in research. Each participant presents their case on three slides:

1. **The case:** research question, what was done, approach, outcome, and the research stages where GenAI was used (stages from GAIDeT, Suchikova et al. 2025).
2. **The setup:** infrastructure, tools and models, data, methods, documentation, time and cost. Concrete details are the point.
3. **Reflection:** barriers, opportunities, a framework or canvas that helped, epistemic authority (how the researcher stays responsible for knowledge claims), a qualitative lens, and new understanding of GenAI practices in research.

`case-template.pptx` has the blank template (slides 1–3) and an illustrative example (slides 4–6). The example, a study of Stack Exchange and other archives in Hugging Face training datasets, is imagined and does not report real results.

## Rebuilding the slides

```sh
pip install python-pptx lxml
python3 build_slides.py
```

## Background

- Suchikova, Y. et al. (2025). Generative AI Delegation Taxonomy (GAIDeT). *Accountability in Research*. https://doi.org/10.1080/08989621.2025.2544331
- Virginia Tech: [Considerations for the Responsible and Ethical Use of AI in Research](https://www.research.vt.edu/research-support/forms-guidance/sirc/considerations-responsible-ethical-use-ai.html)

## Licence

MIT
