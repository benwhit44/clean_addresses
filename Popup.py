import PySimpleGUI as sg

def GUI_POPUP(text, data):
    layout = [
        [sg.Text(text)],
        [sg.Listbox(data, size=(30, 10), key='SELECTED')],
        [sg.Button('OK')],
    ]

    window = sg.Window('Address Cleanup', layout).Finalize()

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'OK':
            break
        else:
            print('OVER')

    window.close()

    # print('[GUI_POPUP] event:', event)
    # print('[GUI_POPUP] values:', values)

    if values and values['SELECTED']:
        return values['SELECTED']