import PySimpleGUI as sg
# import handler
import subprocess
# import sys

sg.theme("BluePurple")


layout = [
    [sg.Frame(layout=[
        [sg.B("配置表", size=(20, 2), key='show_table'), sg.B("生徒コマ登録", size=(20, 2), key='register_student_class'), sg.B("講師コマ登録", size=(20, 2), key='register_teacher_class')]], title='配置')],
    [sg.Frame(layout=[
        [sg.B("生徒登録", size=(20, 2), key='register_new_student'), sg.B("講師登録", size=(20, 2), key='register_new_teacher'), sg.B("Except", size=(20, 2), key='except')]], title='登録')],
    [sg.Frame(layout=[
        [sg.B("配置表印刷", size=(20, 2), key='print_table'), sg.B("生徒スケジュール印刷", size=(20, 2), key='print_student_schedule'), sg.B("print_teacher_schedule", size=(20, 2), key='btn9')]], title='印刷')],
    [sg.Frame(layout=[
        [sg.B("配置表送信", size=(20, 2), key='send_table'), sg.B("初期設定", size=(20, 2), key='initialize_system'), sg.B("バージョン情報", size=(20, 2), key='version_info')]], title='その他')],
    [sg.B("閉じる", size=(20, 2))],
]

window = sg.Window("AST", layout)

while True:  # Event Loop
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED or event == '閉じる':
        break
    else:
        subprocess.run(['python3', f'{event}.py'])


window.close()
