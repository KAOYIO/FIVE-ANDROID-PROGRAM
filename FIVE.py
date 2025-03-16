import os
import pandas as pd
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserIconView, FileChooserListView
from kivy.uix.spinner import Spinner  # 選單功能
from kivy.core.window import Window
from openpyxl import load_workbook

# 設定視窗大小（Windows 適用）
Window.size = (800, 600)

class ExcelEditor(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(orientation='vertical', **kwargs)
        self.data = None  # 儲存 Excel 數據
        self.excel_path = None  # 記錄檔案路徑
        self.text_inputs = []  # 儲存輸入欄位
        self.excel_loaded = False  # 確保 Excel 內容顯示在主視窗內

        # 建立上方選單
        self.create_toolbar()

        # 檔案選擇器 (預設圖示模式)
        self.file_chooser = FileChooserIconView(filters=['*.xlsx'])
        self.file_chooser.bind(on_selection=self.load_excel)
        self.add_widget(self.file_chooser)

    def create_toolbar(self):
        """建立選單列，讓使用者切換檔案選擇器的顯示模式"""
        toolbar = BoxLayout(size_hint_y=0.1)

        # 標題
        title_label = Label(text="Excel 編輯器", size_hint_x=0.7, bold=True)
        toolbar.add_widget(title_label)

        # 下拉選單（切換顯示模式）
        self.view_spinner = Spinner(
            text='選擇顯示模式',
            values=('圖示模式', '清單模式'),
            size_hint_x=0.3
        )
        self.view_spinner.bind(text=self.switch_filechooser_mode)
        toolbar.add_widget(self.view_spinner)

        self.add_widget(toolbar)

    def switch_filechooser_mode(self, spinner, mode):
        """切換檔案選擇器顯示模式"""
        self.remove_widget(self.file_chooser)
        if mode == "圖示模式":
            self.file_chooser = FileChooserIconView(filters=['*.xlsx'])
        else:
            self.file_chooser = FileChooserListView(filters=['*.xlsx'])
        self.file_chooser.bind(on_selection=self.load_excel)
        self.add_widget(self.file_chooser)

    def load_excel(self, instance, selection):
        """讀取 Excel 檔案並顯示內容"""
        if not selection:
            return

        self.excel_path = selection[0]
        try:
            # **解決 Excel 中文亂碼**
            self.data = pd.read_excel(self.excel_path, dtype=str, engine='openpyxl')

            # 讀取成功後，清除檔案選擇器，顯示 Excel 內容
            if not self.excel_loaded:
                self.remove_widget(self.file_chooser)
                self.create_excel_view()
                self.excel_loaded = True

            # 更新 Excel Grid
            self.update_grid()
        except Exception as e:
            print(f"讀取 {self.excel_path} 失敗: {e}")

    def create_excel_view(self):
        """建立 Excel 內容顯示區域，並新增關閉按鈕"""
        self.grid = GridLayout(cols=len(self.data.columns), spacing=5, size_hint_y=None)
        self.grid.bind(minimum_height=self.grid.setter('height'))
        self.add_widget(self.grid)

        # 關閉 Excel 檢視的按鈕
        close_btn_layout = BoxLayout(size_hint_y=0.1)
        close_btn = Button(text="返回", size_hint_x=1)
        close_btn.bind(on_press=self.close_excel_view)
        close_btn_layout.add_widget(close_btn)
        self.add_widget(close_btn_layout)

    def update_grid(self):
        """更新 GridLayout 顯示的 Excel 內容"""
        self.grid.clear_widgets()
        self.text_inputs.clear()

        if self.data is None:
            return

        # **讀取標題**
        for col_name in self.data.columns:
            self.grid.add_widget(Label(text=str(col_name), bold=True))

        # **讀取數據**
        for i, row in self.data.iterrows():
            row_inputs = []
            for j, value in enumerate(row):
                text_input = TextInput(
                    text=str(value),
                    multiline=False,
                    background_color=(1, 1, 1, 1),  # **解決紅點問題**
                    foreground_color=(0, 0, 0, 1),
                    use_bubble=False,  # **禁用右鍵菜單**
                    use_handles=False  # **禁用拖曳選取**
                )
                text_input.bind(focus=self.remove_red_dot)  # 修正紅點問題
                row_inputs.append(text_input)
                self.grid.add_widget(text_input)
            self.text_inputs.append(row_inputs)

    def remove_red_dot(self, instance, value):
        """修正 Kivy 在 Windows 上 TextInput 產生紅點的問題"""
        if value:
            instance.background_color = (1, 1, 1, 1)  # 保持白色背景
        else:
            instance.background_color = (1, 1, 1, 1)

    def close_excel_view(self, instance):
        """關閉 Excel 內容顯示，返回檔案選擇視窗"""
        self.clear_widgets()
        self.create_toolbar()
        self.file_chooser = FileChooserIconView(filters=['*.xlsx'])
        self.file_chooser.bind(on_selection=self.load_excel)
        self.add_widget(self.file_chooser)
        self.excel_loaded = False  # 重置狀態

class ExcelApp(App):
    def build(self):
        return ExcelEditor()

if __name__ == '__main__':
    ExcelApp().run()
｛
