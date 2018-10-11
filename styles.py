##Openpyxl Styles

from openpyxl.styles.borders import Border, Side

#border

thin_bottom_border = Border(left=Side(style=None), 
                     right=Side(style=None), 
                     top=Side(style=None), 
                     bottom=Side(style='thin'))

double_bottom_border = Border(left=Side(style=None), 
                     right=Side(style=None), 
                     top=Side(style=None), 
                     bottom=Side(style='double'))