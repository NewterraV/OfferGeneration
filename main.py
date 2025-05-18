import flet as ft

from src.ui_logic import OfferGeneratorApp


def main(page: ft.Page):
    page.window.width = 560
    page.window.height = 900
    page.window.resizable = False
    page.theme_mode = ft.ThemeMode.DARK
    page.window.icon = "data/icon.png"
    OfferGeneratorApp(page)


if __name__ == '__main__':
    ft.app(target=main)