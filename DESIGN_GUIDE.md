# DESIGN_GUIDE

Updated: 2026-08-18

## Interface Direction

The application should feel like a compact work tool inside Bitrix24, not a marketing page.

## Layout

- First screen is the actual spreadsheet widget.
- Header is short and functional: app name, deal context, field reload action.
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
- Field picker appears next to the selected cell control.
- Field list shows human-readable field title and technical field ID.
- Mobile layout stacks the topbar and keeps the table horizontally scrollable.
