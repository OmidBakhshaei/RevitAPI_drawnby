# pylint: disable=import-error,invalid-name
import ctypes
import sys
from pyrevit import revit
from pyrevit import forms
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, Transaction
from rpw.ui.forms import TextInput, TaskDialog, CommandLink

dialog = TaskDialog("Add or Overwrite", commands=None, buttons=[
                    'Yes', 'No'], title='Drawn By', content='Are you the first person adding your initials to "Drawn By" parameter on the sheets?', title_prefix=False, show_close=True, footer='', expanded_content='', verification_text='')
answer = dialog.show()

if answer == None:
    sys.exit()

if answer == 'No':
    commands = [CommandLink('Comma', return_value='comma'),
                CommandLink('Space', return_value='space')]
    dialog2 = TaskDialog("Separator", commands=commands, buttons=None, title='Drawn By', content='Should a comma or a space separate your initials from the others?',
                         title_prefix=False, show_close=True, footer='', expanded_content='', verification_text='')
    comma_or_space = dialog2.show()

try:
    if comma_or_space == None:
        sys.exit()
except:
    pass


# Get user's name
def get_display_name():
    GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
    NameDisplay = 3

    size = ctypes.pointer(ctypes.c_ulong(0))
    GetUserNameEx(NameDisplay, None, size)

    nameBuffer = ctypes.create_unicode_buffer(size.contents.value)
    GetUserNameEx(NameDisplay, nameBuffer, size)
    return nameBuffer.value


initials = get_display_name().split(
    ' ')[0][0]+get_display_name().split(' ')[1][0]
user_initials = TextInput('Drawn by', default=initials,
                          description="Please enter your initials:")

uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
if uidoc:
    doc = uiapp.ActiveUIDocument.Document

sel_sheets = forms.select_sheets(title='Select Titleblocks')

selected_sheet_names = []
parameter_to_change = title_block.LookupParameter("Drawn By")
parameter_name = "Initials"


if sel_sheets:
    for sheet in sel_sheets:
        selected_sheet_names.append(sheet.LookupParameter("Sheet Number").AsString(
        ) + " - " + sheet.LookupParameter("Sheet Name").AsString())

title_blocks = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_TitleBlocks)\
    .WhereElementIsNotElementType()\
    .ToElements()

t = Transaction(doc, "Change params")

t.Start()

param_error = []
overwritten_sheets = []
overwritten_initials = []

if answer == "Yes":
    # Update the parameters
    for title_block in title_blocks:
        if title_block.LookupParameter(parameter_to_change) and title_block.LookupParameter("Sheet Number").AsString()+" - "+title_block.LookupParameter("Sheet Name").AsString() in selected_sheet_names:
            existing_value = title_block.LookupParameter(
                parameter_to_change).AsString()
            if existing_value != "Author" and existing_value != "" and existing_value != user_initials:
                overwritten_sheets.append(title_block.LookupParameter("Sheet Number").AsString(
                ) + " - " + title_block.LookupParameter("Sheet Name").AsString())
                overwritten_initials.append(existing_value)
            title_block.LookupParameter(parameter_to_change).Set(user_initials)
        # Collect the sheets that do not have the parameter and failed to be updated
        elif not title_block.LookupParameter(parameter_to_change) and title_block.LookupParameter("Sheet Number").AsString()+" - "+title_block.LookupParameter("Sheet Name").AsString() in selected_sheet_names:
            param_error.append("{0} - {1}".format(title_block.LookupParameter(
                "Sheet Number").AsString(), title_block.LookupParameter("Sheet Name").AsString()))
elif answer == "No":
    for title_block in title_blocks:
        if title_block.LookupParameter(parameter_to_change) and title_block.LookupParameter("Sheet Number").AsString()+" - "+title_block.LookupParameter("Sheet Name").AsString() in selected_sheet_names:
            if comma_or_space == 'comma':
                others_initials = title_block.LookupParameter(
                    parameter_to_change).AsString() + ", "
            elif comma_or_space == 'space':
                others_initials = title_block.LookupParameter(
                    parameter_to_change).AsString() + " "
            if "Author" in others_initials:
                others_initials = ""
            new_drawn_by = others_initials + user_initials
            title_block.LookupParameter(parameter_to_change).Set(new_drawn_by)
        # Collect the sheets that do not have the parameter and failed to be updated
        elif not title_block.LookupParameter(parameter_to_change) and title_block.LookupParameter("Sheet Number").AsString()+" - "+title_block.LookupParameter("Sheet Name").AsString() in selected_sheet_names:
            param_error.append("{0} - {1}".format(title_block.LookupParameter(
                "Sheet Number").AsString(), title_block.LookupParameter("Sheet Name").AsString()))
else:
    pass

t.Commit()

# Print the sheets which returned errors
if param_error:
    print('\nATTENTION !!!')
    print('\nThe following sheets returned errors (They might miss the "Drawn By" parameter or have different naming convention) :'.format(
        parameter_name).upper())
    for i in param_error:
        print(i)
    print('\n'+'-' * 100)

# Print the sheets which were successfully updated
if sel_sheets:
    print('\nYour {0} ({1}) were successfully added to the following sheets:'.format(
        parameter_name, user_initials))
    for sheet in sel_sheets:
        snum = sheet.Parameter[BuiltInParameter.SHEET_NUMBER]\
            .AsString()
        sname = sheet.Parameter[BuiltInParameter.SHEET_NAME]\
            .AsString()
        if '{0} - {1}'.format(snum, sname) not in param_error and user_initials in sheet.LookupParameter(parameter_to_change).AsString():
            print('{0} - {1}'.format(snum, sname))
# Let the user know that no change was made
else:
    forms.alert("No sheets were selected, {0}!".format(
        get_display_name().split(' ')[0]), title='Alert!', ok=True)

if overwritten_sheets:
    print('\nAttention !!!'.upper())
    for i in range(len(overwritten_sheets)):
        print('\nYou removed "{0}" from the sheet "{1}" "Drawn By" parameter'.format(
            overwritten_initials[i], overwritten_sheets[i]))
    print('\nIf you have accidentally removed someone elses initials from the sheets, please undo (Ctrl+Z), re-run the script, and answer "No" to the question "Are you the first person adding your initials to "Drawn By" on the sheets?"'.upper())
