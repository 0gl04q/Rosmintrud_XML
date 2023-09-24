import os
import sys

import flet as ft

import functions


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def main(page: ft.Page):
    wb = ''

    def find_option(option_name):
        for option in dropdown.options:
            if option_name == option.key:
                return option
        return None

    def show_banner_click(e, exceptions):
        page.banner.content = ft.Text(exceptions)
        page.banner.open = True
        page.update()

    def close_banner(e):
        page.banner.open = False
        page.update()

    def open_dlg(e, text):
        page.dialog = ft.AlertDialog(title=ft.Text(text))
        page.dialog.open = True
        page.update()

    def get_list_protocol_button(e):
        nonlocal wb
        wb = functions.get_workbook()

        # TODO: Добавить ожидание

        if isinstance(wb, str):
            show_banner_click(e, wb)
        else:
            dropdown.options = [ft.dropdown.Option(protocol) for protocol in functions.get_list_protocol(wb)]
            open_dlg(e, "Список протоколов сформирован!")
        page.update()

    def create_xml_file(e):
        option = find_option(dropdown.value)
        if option is not None:
            end = functions.create_xml(option.key, wb)
            if isinstance(end, str):
                show_banner_click(e, end)
            else:
                open_dlg(e, f'Файл {option.key}.xml успешно сформирован!')
                dropdown.options.remove(option)
            page.update()

    def update_file(e: ft.FilePickerResultEvent):
        if e.files:
            open_dlg(e, functions.data_update(e.files[0].path, wb))
        page.update()

    page.banner = ft.Banner(
        bgcolor=ft.colors.AMBER_100,
        leading=ft.Icon(ft.icons.WARNING_AMBER_ROUNDED, color=ft.colors.AMBER, size=40),
        actions=[ft.TextButton("Закрыть", on_click=close_banner)],
    )

    pick_files_dialog = ft.FilePicker(on_result=update_file)

    dropdown = ft.Dropdown(
        label="Список протоколов",
        hint_text="Выберите протокол",
    )

    page.overlay.append(pick_files_dialog)
    page.window_width = 400
    page.window_height = 400
    page.title = 'Create XML'
    page.add(
        ft.ResponsiveRow(
            [
                dropdown,
                ft.ElevatedButton("Получить список протоколов",
                                  icon=ft.icons.GET_APP_OUTLINED,
                                  on_click=get_list_protocol_button,
                                  col={"md": 4}),
                ft.ElevatedButton("Сформировать XML",
                                  icon=ft.icons.MESSENGER_OUTLINE,
                                  on_click=create_xml_file,
                                  col={"md": 4}),
                ft.ElevatedButton("Обратная загрузка",
                                  icon=ft.icons.UPLOAD_FILE,
                                  on_click=pick_files_dialog.pick_files,
                                  col={"md": 4}),
            ],
        ),
    )


if __name__ == '__main__':
    ft.app(target=main)
