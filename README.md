# Student Record Manager

A local-first browser tool for controlled Moodle → Obsidian student-record imports. This tool is specific to my workflow and an upgrade from my previous [XLS2MD](https://github.com/haakmal/xls2md) project.

## Workflow

> [!important]
> This tool follows my specific Obsidian library setup, if you wish to utilise it for your own purposes then you would need to rework a fair but unless you wish to retain my note structure.

1. Select the Moodle XLSX.
2. Enter the course, year, term, and student level (UG or PG).
3. Select the Obsidian student-record directory.
4. Select the Obsidian tutor-record directory.
5. Scan the Moodle records.
6. Assign a tutor to each Moodle group.
7. Build the review queue.
8. Preview and individually apply proposed student-record changes.

### Current Obsidian conventions

As mentioned, the application deliberately preserves existing student-note schema, including the current frontmatter/property format and course syntax. **Editing these portions would need remapping in the code.**

```yaml
---
ID: z1234567
aliases:
  - John Doe
email: j.doe@student.unsw.edu.au
tags: DDES1150/2026/T3
level: UG
mentor: false
problem: false
honours: false
student: true
---
```

Course records remain:

```markdown
## [[DDES1150]]

Tutor:: [[Tutor Name]]

> [!warning] Warning

> [!notes] Notes
```

## Running locally

For cleanest and least obstructive experience, use a Chromium-based browser such as Chrome or Edge for the folder read/write APIs.

From the project directory in terminal run:

```bash
python3 -m http.server 8080
```

Open `http://localhost:8080/` to access the web UI.

## Demo data

The repository includes a dummy Moodle export and a small demo vault under `demo-vault/` so the workflow can be tested without touching a real Obsidian vault.

## AI Collaboration

This project was designed and developed by Dr Haider Ali Akmal in collaboration with an AI-supported design and development partner. Naming a specific model would be insufficient and inaccurate as multiple iterations may have been analysed and collaborated on with different models.

AI was used at key points in this project as a collaborative tool for activities including code development and debugging, interface iteration, content refinement, and critical discussion of design decisions. The concept, pedagogical direction, design requirements, evaluation, and final decision-making remain the work and responsibility of the project author.

This acknowledgement reflects a commitment to transparency around AI-assisted creative and technical practice, and an interest in exploring human–AI collaboration as an evolving mode of design practice.
