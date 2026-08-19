# OfficeCLI Benchmarks

This directory contains evaluation systems that use OfficeCLI's structured
inspection and rendering output.

## Office Reward

[`office-reward/`](office-reward/) is a fine-grained PPTX, DOCX, and XLSX
scoring experiment and browser-based human annotation workbench. Its checked-in
evidence covers 54 rendered Office units, 2,160 direct model subquestion scores,
and 1,080 explicit Content Accuracy abstentions.

The benchmark is independent from the OfficeCLI binary build. Run and test it
from its own directory with Node.js 22 or newer.
