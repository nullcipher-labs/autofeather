from comtypes.client import GetActiveObject
import win32com.client
from PIL import Image


def create_feathered_copy(file_path, feather, prefix, postfix, use_percentage):
    """a method that takes in a file path of an image file, amount of feather, prefix and postfix,
    and creates a png copy of the original image, feathered by the requested amount of pixels, and named
    using the prefix and postfix

    :param file_path: str, path to the image file
    :param feather: int, amount of pixels to fade from the frame
    :param prefix: str, string to add at the beginning of the new file name
    :param postfix: str, string to add at the end of the new file name
    :param use_percentage: bool, determines if the number in feather is used as plain pixels or percentage of the
    smaller dimension of the image
    :return: str, a string confirming the new file has been created
    """
    # calculates pixels for border by percentage of the smaller dimension of the image
    if use_percentage:
        with Image.open(file_path) as img:
            w, h = img.size
            feather = min(w, h) * (feather/100)

    # creates the new file name and new png file path
    file_name = file_path.split('\\')[-1].split('.')[0]
    new_name = prefix + file_name + postfix
    png_path = '\\'.join(file_path.split('\\')[:-1]) + f'\\{new_name}.png'

    # setting up params for Photoshop
    ps_inches = 2
    ps_pixels = 1
    ps_text_layer = 2
    ps_replace_selection = 1

    app = GetActiveObject("Photoshop.Application")

    start_ruler_units = app.Preferences.RulerUnits
    if start_ruler_units is not ps_inches:
        app.Preferences.RulerUnits = ps_inches

    # opening Photoshop
    src_doc = app.Open(file_path)

    if src_doc.ActiveLayer.Kind != ps_text_layer:
        # creating the feathered selection
        x2 = src_doc.Width * src_doc.Resolution
        y2 = src_doc.Height * src_doc.Resolution

        sel_area = ((0, 0), (x2, 0), (x2, y2), (0, y2))
        src_doc.Selection.Select(sel_area, ps_replace_selection, 0, False)

        src_doc.Selection.Copy()

        app.Preferences.RulerUnits = ps_pixels
        paste_doc = app.Documents.Add(x2, y2, src_doc.Resolution, new_name)
        paste_doc.Paste()

        layer_ref = paste_doc.ArtLayers["Background"]
        layer_ref.Delete()

        src_doc.Close()

        # deletes the selection 20 times (deletes all there is to delete,
        # one deletion deletes the selection only partly, 20 is a sufficiently large number of deletions
        # to get all the selection deleted for every photo)
        for i in range(20):
            paste_doc.Selection.Select(sel_area, ps_replace_selection, feather, False)
            paste_doc.Selection.Invert()
            paste_doc.Selection.Cut()
    else:
        print("You cannot copy from a text layer")

    if start_ruler_units != app.Preferences.RulerUnits:
        app.Preferences.RulerUnits = start_ruler_units

    # using win32com to save results as png
    ps_app = win32com.client.Dispatch("Photoshop.Application")
    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 13
    options.PNG8 = False
    doc = ps_app.Application.ActiveDocument
    doc.Export(ExportIn=png_path, ExportAs=2, Options=options)
    ps_app.ActiveDocument.Close(2)

    return f'Saved {png_path}'
