# sheetjs

Apply a simple spreadsheet UI to your table tag as a jQuery-plugin.

Just write **'sheetjs'** class into your **'table'** tag!

<img src="doc/sample.gif">


## Usage
1. Import jQuery(1.x, 2.x or 3.x) and Material Icons from CDN
```html
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<link rel="stylesheet" href="https://fonts.googleapis.com/icon?family=Material+Icons">
```
2. Import jquery.sheetjs.css and jquery.sheetjs.js
```html
<link rel="stylesheet" href="jquery.sheetjs.css" />
<script src="jquery.sheetjs.js"></script>
```
3. Set "sheetjs" class to your table tag
```html
<table class="sheetjs">
  <caption>Table name(optional)</caption>
  <thead>
    <tr>
      <th>A</th>
      <th>B</th>
      <th>..</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>a</td>
      <td>b</td>
      <td>..</td>
    </tr>
    ..
  </tbody>
  <tfoot>
    (optional)
  </tfoot>
</table>
```

That's it!


Please see the [sample.html](sample.html) .

## Specifications

- Prerequisites
  - Table tag is complete. Use thead, tbody, tfoot and caption properly.
  - Works on both IE7 and Chrome.
  
- Functions
  - Cell selection
    - Click the cell to select it.
    - Click with Shift-key, the rectangle will be selected.
    - If another area is selected, the selection will be removed.
  - Line selection
    - Click 'th' at the beginning of a line in tbody, you can also select the line
    - Clicking on other element on th is not work
  - Column selection
    - Click on a colum 'thead th' to select all the columns with thead, tbody and tfoot.
    - 'th' only. 'td' does not react.
  - Value copy
    - Selecting cell text can be copy into your clipboard with Copy command, 'Ctrl + c' or right-click.
    - The copied value is in TSV (tab separated) text format and can be pasted into Excel as it is.
    - The text does not include INPUT or A tag link information. Extract only pure text
  - Filter
    - Select a cell and right-click to Filter to filter the entire table with the selected value
  - Line count display
    - The number of displayed tbody columns is automatically counted and displayed on the screen. If you have declared a caption tag.
  - Total value display
    - Sum the value of the selected cell and display it. If you have declared a caption tag.
    - Take out what looks like a number and add it up. For example, if you select 1,000 yen and A-100, 1000+ (-100) will be 900.
  - Header fixed
    - Follow only the header title as the window scrolls
  
- Functional restrictions in IE
  - No header fixed
  - No icons in IE7
