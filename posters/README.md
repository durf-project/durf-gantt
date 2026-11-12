# Theme posters for the DURF Federation Day

One A3 landscape poster per theme for the theme carousel at the DURF Federation Day
(18 November 2026). Each poster lies on a table; the theme lead explains the goal of the
theme, the activities that lead to it, what we expect institutions to contribute and when.
People at the table write down which in-kind contribution their institution could make, to
which activity, and who in their organisation could do it.

Online: https://durf-project.github.io/durf-gantt/posters/ (everything), `?theme=0` for the opening poster, or `?theme=1` to `?theme=6`.

The **opening poster** shows the six themes as pieces of one puzzle, with how the carousel works and the theme leads. Each theme poster carries its own piece of that puzzle (in the SURF puzzle style) with the theme number in it. The pieces are the DURF puzzle from the design system (`assets/icons/puzzle-set/`, class `.durf-piece`).

## Where the content comes from

| On the poster | Source |
| --- | --- |
| Theme name, goal, partners | `Themes` sheet of `data/durf-roadmap.xlsx` |
| Theme lead | the DURF design system (`durf-themes.json`), the current leads |
| Activities, output, months, milestones | `Activities` sheet: `activity`, `outcomes`, `start_month`, `end_month`, `milestone_months` |
| What we ask from institutions | `Activities` column `needs: in-kind contributions` |
| When (during / after / both) | `Activities` column `when: ...` |
| Event line, the question, the theme icons, the texts of the opening poster | `POSTER` at the top of the script in `index.html` |

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
on the page. When a theme would come out smaller than 8.5 pt (`MIN_PT` in `index.html`), it is
spread over two posters ("part 1 of 2", "part 2 of 2"), split where both sheets get the
largest text. Today that happens for themes 1 and 2, so a full print is 9 pages with the opening poster. Shorter in-kind
texts in the workbook give larger text (the size used is in the `data-font-size` attribute of
each poster).

The rows are banded in a light and a lighter tint of the theme colour. The theme's puzzle piece carries
the theme number; the footer uses the DURF logo without
ring text.

Print from Chrome or Edge with **Print or save as PDF**: paper size A3, landscape, margins
*None*, *Background graphics* on. The topbar, the toolbar and the feedback button are left out of
the print.

## Files

- `index.html`: the posters (DURF design system: colours, fonts, puzzle pieces).
- `qr-roadmap.svg`, `qr-home.svg`: QR codes to the roadmap and the DURF home page.
