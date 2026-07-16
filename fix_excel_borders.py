def style_range(ws, cell_range, border):
    """
    Apply styles to a range of cells as if they were a single cell.
    """
    from openpyxl.styles import Border, Side
    
    top = Border(top=border.top)
    left = Border(left=border.left)
    right = Border(right=border.right)
    bottom = Border(bottom=border.bottom)
    
    first_cell = ws[cell_range.split(":")[0]]
    rows = ws[cell_range]
    
    for cell in rows[0]:
        cell.border = cell.border + top
    for cell in rows[-1]:
        cell.border = cell.border + bottom
    
    for row in rows:
        l = row[0]
        r = row[-1]
        l.border = l.border + left
        r.border = r.border + right
