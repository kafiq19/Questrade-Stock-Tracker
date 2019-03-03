##Openpyxl Styles

from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font, PatternFill

#border

thin_bottom_border = Border(left=Side(style=None), 
                     right=Side(style=None), 
                     top=Side(style=None), 
                     bottom=Side(style='thin'))

double_bottom_border = Border(left=Side(style=None), 
                     right=Side(style=None), 
                     top=Side(style=None), 
                     bottom=Side(style='double'))

bold_font = Font(bold=True)

fill_pattern = PatternFill(bgColor="FFC7CE", fill_type = "solid")