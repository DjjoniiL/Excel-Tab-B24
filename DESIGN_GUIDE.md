# DESIGN_GUIDE

Updated: 2026-08-18

## Interface Direction

The application should feel like a compact work tool inside Bitrix24, not a marketing page.

## Layout

- First screen is the actual spreadsheet widget.
- Header is short and functional: app name, deal context, field reload action.
- Toolbar includes compact table commands such as selecting filled cells and selected-range actions.
- Toolbar includes export as a direct table action.
- Table area is the dominant surface.
- Row add control sits below the table.
- Column add control sits to the right of the table.

## Visual System

- Base background: light neutral gray.
- Main table surface: white with clear grid lines.
- Accent: blue for selected cells and primary actions.
- Status elements: compact bordered pills.
- Border radius is limited to 6px.

## Interaction Notes

- Cell controls stay quiet until focus.
- Selection actions appear only after at least one cell is selected.
- Field picker appears next to the selected cell control.
- Field picker can be dismissed with a visible close button, outside click, Escape, or repeated click on the same picker control.
- Field list shows human-readable field title and technical field ID.
- Mobile layout stacks the topbar and keeps the table horizontally scrollable.
