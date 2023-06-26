import PySimpleGUI as sg
import pandas as pd

# Add somtre color to the window

sg.theme('DarkBrown2')   # Add a fdftoucdfdgdgfdfh of colorssdfdfddf

EXCEL_FILE = 'test2.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text("Please fill out the following fields:")],
    [sg.Text("Name", size=(15, 1)), sg.InputText(key="Name")],
    [sg.Text("Favorite Colour", size=(15, 1)), sg.Combo(["Red", "Green", "Blue"], key="Favorite Color")],
    [sg.Text("Favorite Animal", size=(15, 1)), sg.Combo(["Dog", "Cat", "Rabbit"], key="Favorite Animal")],
    [sg.Text("Favorite Movie", size=(15, 1)), sg.Combo(["Star Wars", "Harry Potter", "Lord of the Rings"], key="Favorite Movie")],
    [sg.Text("I speak", size=(15, 1)), sg.Checkbox("English", key="English"),sg.Checkbox("French", key="French"),sg.Checkbox("Spanish", key="Spanish")],
    [sg.Text("No. of Children", size=(15,1)), sg.Spin([i for i in range(1, 16)],
                                                      initial_value=0, key="Children")],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window("Form", layout)


def clear_input():
    for key in values:
        window[key]("")
    return None


while True:
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    if event == 'Clear':
        clear_input()
    if event == "Submit":
        print(values)
        print(event)
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup("Form submitted successfully!")
        clear_input()



window.close()