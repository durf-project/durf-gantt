# Theme posters for the DURF Federation Day

One A3 landscape poster per theme for the theme carousel at the DURF Federation Day
(18 November 2026). Each poster lies on a table; the theme lead explains the goal of the
theme, the activities that lead to it, what we expect institutions to contribute and when.
People at the table write down which in-kind contribution their institution could make, to
which activity, and who in their organisation could do it.

Online: https://durf-project.github.io/durf-gantt/posters/ (all six) or `?theme=1` to `?theme=6`.

## Where the content comes from

| On the poster | Source |
| --- | --- |
| Theme name, goal, partners | `Themes` sheet of `data/durf-roadmap.xlsx` |
| Theme lead | the DURF design system (`durf-themes.json`), the current leads |
| Activities, output, months, milestones | `Activities` sheet: `activity`, `outcomes`, `start_month`, `end_month`, `milestone_months` |
| What we ask from institutions | `Activities` column `needs: in-kind contributions` |
| When (during / after / both) | `Activities` column `when: ...` |
| Event line, the question, the theme icons | `POSTER` at the top of the script in `index.html` |

So edit the workbook in the Nextcloud team folder; the sync workflow rebuilds
`data/roadmap.json`, and the posters follow. `notes` and the other private columns are not shown.

### Writing the in-kind column

Each contribution on its own line, starting with its level:

    essential: ... to have your library director sign the position paper, before M12. [during]
    optional: ... to have someone take part in the NaMeCo meetings. [both]

`Optional:` or `Essential:` on a line of its own sets the level for the `- ` lines below it.
`[during]`, `[after]` or `[both]` at the end becomes a small label: during the project,
after the project, or during and after.

## Size and printing

The posters are A3 landscape (420 × 297 mm). The activity table shrinks its font until it fits
on the page, so a theme with many activities or long texts gets smaller text: shorten the
in-kind texts in the workbook when that gets too small to read (the font size used is in the
`data-font-size` attribute of each poster).

Print from Chrome or Edge with **Print or save as PDF**: paper size A3, landscape, margins
*None*, *Background graphics* on. The topbar, the toolbar and the feedback button are left out of
the print.

## Files

- `index.html`: the posters (DURF design system: colours, fonts, puzzle pieces, pictograms).
- `qr-roadmap.svg`, `qr-home.svg`: QR codes to the roadmap and the DURF home page.
