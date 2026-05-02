"""設定ダイアログ."""

from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from config.config_manager import ConfigManager


class SettingsDialog(QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("設定")
        self.setMinimumWidth(480)
        self._mgr = ConfigManager()
        self._cfg = self._mgr.config
        self._build_ui()
        self._load_into_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        tabs = QTabWidget()
        tabs.addTab(self._build_capture_tab(), "キャプチャ")
        tabs.addTab(self._build_navigation_tab(), "ページ送り")
        tabs.addTab(self._build_hotkey_tab(), "ホットキー")
        tabs.addTab(self._build_storage_tab(), "保存")
        layout.addWidget(tabs)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    # ---------------- tabs ----------------
    def _build_capture_tab(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)

        self.spin_delay = QDoubleSpinBox()
        self.spin_delay.setRange(0.1, 10.0)
        self.spin_delay.setSingleStep(0.1)
        self.spin_delay.setSuffix(" 秒")
        form.addRow("ページ送り後の待機:", self.spin_delay)

        self.spin_pre_delay = QDoubleSpinBox()
        self.spin_pre_delay.setRange(0.0, 5.0)
        self.spin_pre_delay.setSingleStep(0.05)
        self.spin_pre_delay.setSuffix(" 秒")
        form.addRow("初回フォーカス後の遅延:", self.spin_pre_delay)

        self.cmb_format = QComboBox()
        self.cmb_format.addItems(["png", "jpeg", "webp"])
        form.addRow("保存フォーマット:", self.cmb_format)

        self.spin_jpeg_q = QSpinBox()
        self.spin_jpeg_q.setRange(1, 100)
        form.addRow("JPEG 品質:", self.spin_jpeg_q)

        self.spin_webp_q = QSpinBox()
        self.spin_webp_q.setRange(1, 100)
        form.addRow("WebP 品質:", self.spin_webp_q)

        self.chk_skip_dup = QCheckBox("重複ページをスキップ")
        form.addRow(self.chk_skip_dup)

        self.spin_dup_thresh = QSpinBox()
        self.spin_dup_thresh.setRange(0, 64)
        self.spin_dup_thresh.setToolTip(
            "この距離以下を「同じページ」とみなし保存をスキップします（hash_size=16 の場合 256bit 中の差分ビット数）。\n"
            "大きすぎると隣接ページが重複扱いになり全スキップされます。\n"
            "推奨: 3（完全一致に近いものだけスキップ）。ページが飛ぶ場合は 2 以下にしてください。"
        )
        form.addRow("pHash 重複閾値（スキップ送り用）:", self.spin_dup_thresh)

        self.spin_last_page_dist = QSpinBox()
        self.spin_last_page_dist.setRange(0, 32)
        self.spin_last_page_dist.setToolTip(
            "このハミング距離以下のときだけ「最終ページ用」の連続カウントに含めます。"
            "大きいと隣接ページの誤検知で止まりやすく、小さいと本当の最終ページで止まりにくいです。"
        )
        form.addRow("最終ページ判定の厳密距離上限:", self.spin_last_page_dist)

        self.chk_auto_stop = QCheckBox("最終ページ検知で自動停止")
        form.addRow(self.chk_auto_stop)

        self.spin_consec_dupes = QSpinBox()
        self.spin_consec_dupes.setRange(1, 10)
        form.addRow("最終ページ判定 連続重複回数:", self.spin_consec_dupes)

        self.spin_min_cap_last = QSpinBox()
        self.spin_min_cap_last.setRange(0, 1000)
        self.spin_min_cap_last.setToolTip(
            "この枚数以上を保存したあとだけ、連続重複で最終ページとみなします。"
            "推奨 5 以上。アプリ内部でさらに下限 5 がかかります。"
        )
        form.addRow("最終ページ判定までの最低保存枚数:", self.spin_min_cap_last)

        self.spin_max_dup_skips = QSpinBox()
        self.spin_max_dup_skips.setRange(5, 5000)
        self.spin_max_dup_skips.setToolTip(
            "保存せず重複スキップだけがこの回数続いたらエラー停止します。"
            "ページ送りが Kindle に届いていないときの無限ループ防止です。"
        )
        form.addRow("重複スキップ連続の上限(保存なし):", self.spin_max_dup_skips)

        self.chk_printwindow = QCheckBox("ウィンドウ内容のみ取得（PrintWindow・重なり除外）")
        self.chk_printwindow.setToolTip(
            "オン: Kindle ウィンドウの描画だけを取り、Cursor など他ウィンドウが重なっても写しません。\n"
            "オフ: 画面の矩形をそのまま取ります（従来どおり）。\n"
            "※ 真っ黒になる環境ではオフを試してください。"
        )
        form.addRow(self.chk_printwindow)

        return w

    def _build_navigation_tab(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)

        self.cmb_nav_method = QComboBox()
        self.cmb_nav_method.addItems(["key", "click"])
        form.addRow("ページ送り方式:", self.cmb_nav_method)

        self.edit_nav_key = QLineEdit()
        form.addRow("送りキー:", self.edit_nav_key)

        self.cmb_key_delivery = QComboBox()
        self.cmb_key_delivery.addItems(
            ["sendinput", "wheel_pyautogui", "wheel_postmessage", "postmessage", "pyautogui"]
        )
        self.cmb_key_delivery.setToolTip(
            "sendinput: Page Down 等（本アプリ最小化＋既定 pagedown が最も確実）。\n"
            "wheel_*: ホイールは 1 ノッチ=1 ページ。進行が逆なら下の「符号」を反転。\n"
            "postmessage / pyautogui: 補助。"
        )
        form.addRow("キー送り方式:", self.cmb_key_delivery)

        self.spin_wheel_notches = QSpinBox()
        self.spin_wheel_notches.setRange(1, 10)
        self.spin_wheel_notches.setToolTip(
            "ホイール方式のみ有効。3 にすると 1 回で最大 3 ページ動き、表紙に戻る原因になります。通常は 1。"
        )
        form.addRow("ホイールノッチ数(1推奨):", self.spin_wheel_notches)

        self.cmb_wheel_sign = QComboBox()
        self.cmb_wheel_sign.addItem("-1（次ページ・多くの環境）", -1)
        self.cmb_wheel_sign.addItem("+1（進行が逆のとき）", 1)
        form.addRow("ホイール符号:", self.cmb_wheel_sign)

        self.chk_invert_nav = QCheckBox("ページ送り方向を反転（キー＋ホイール）")
        self.chk_invert_nav.setToolTip(
            "Page↔PageUp、左右・上下矢印の入れ替えと、ホイール符号の反転をまとめて行います。"
        )
        form.addRow(self.chk_invert_nav)

        self.spin_click_x = QSpinBox()
        self.spin_click_x.setRange(-2000, 2000)
        form.addRow("クリックXオフセット:", self.spin_click_x)

        self.spin_click_y = QSpinBox()
        self.spin_click_y.setRange(-2000, 2000)
        form.addRow("クリックYオフセット:", self.spin_click_y)

        self.chk_refocus = QCheckBox("操作前にウィンドウを前面化")
        form.addRow(self.chk_refocus)

        return w

    def _build_hotkey_tab(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)

        self.edit_hk_start = QLineEdit()
        form.addRow("開始/停止:", self.edit_hk_start)

        self.edit_hk_pause = QLineEdit()
        form.addRow("一時停止:", self.edit_hk_pause)

        self.edit_hk_emerg = QLineEdit()
        form.addRow("緊急停止:", self.edit_hk_emerg)

        note = QLabel(
            "例: f9 / ctrl+shift+s / alt+space\n"
            "ホットキーは保存後、再起動で反映されます"
        )
        note.setStyleSheet("color: #888;")
        form.addRow(note)
        return w

    def _build_storage_tab(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)

        self.edit_subfolder = QLineEdit()
        form.addRow("サブフォルダテンプレ:", self.edit_subfolder)

        self.edit_filetpl = QLineEdit()
        form.addRow("ファイル名テンプレ:", self.edit_filetpl)

        self.chk_resume = QCheckBox("既存ファイル番号から再開する")
        form.addRow(self.chk_resume)

        self.chk_auto_pdf = QCheckBox("セッション終了時に PDF を自動生成")
        self.chk_auto_pdf.setToolTip(
            "取得した連番画像を 1 つの PDF にまとめ、セッションフォルダに保存します。"
        )
        form.addRow(self.chk_auto_pdf)

        self.edit_pdf_name = QLineEdit()
        self.edit_pdf_name.setPlaceholderText("book.pdf")
        form.addRow("PDF ファイル名:", self.edit_pdf_name)

        return w

    # ---------------- load/save ----------------
    def _load_into_ui(self) -> None:
        c = self._cfg
        self.spin_delay.setValue(c.capture.page_delay_sec)
        self.spin_pre_delay.setValue(c.capture.pre_capture_delay_sec)
        self.cmb_format.setCurrentText(c.capture.save_format)
        self.spin_jpeg_q.setValue(c.capture.jpeg_quality)
        self.spin_webp_q.setValue(c.capture.webp_quality)
        self.chk_skip_dup.setChecked(c.capture.skip_duplicates)
        self.spin_dup_thresh.setValue(c.capture.duplicate_threshold)
        self.spin_last_page_dist.setValue(c.capture.last_page_max_hash_distance)
        self.chk_auto_stop.setChecked(c.capture.auto_stop_on_last_page)
        self.spin_consec_dupes.setValue(c.capture.last_page_consecutive_dupes)
        self.spin_min_cap_last.setValue(c.capture.min_captures_before_last_page_stop)
        self.spin_max_dup_skips.setValue(c.capture.max_duplicate_skips_without_save)
        self.chk_printwindow.setChecked(bool(getattr(c.capture, "use_printwindow", True)))

        self.cmb_nav_method.setCurrentText(c.navigation.method)
        self.edit_nav_key.setText(c.navigation.key)
        kd = getattr(c.navigation, "key_delivery", "sendinput") or "sendinput"
        i = self.cmb_key_delivery.findText(kd)
        self.cmb_key_delivery.setCurrentIndex(max(0, i))
        self.spin_wheel_notches.setValue(getattr(c.navigation, "wheel_scroll_notches", 1))
        sign = int(getattr(c.navigation, "wheel_scroll_sign", -1))
        for idx in range(self.cmb_wheel_sign.count()):
            if self.cmb_wheel_sign.itemData(idx) == sign:
                self.cmb_wheel_sign.setCurrentIndex(idx)
                break
        else:
            self.cmb_wheel_sign.setCurrentIndex(0)
        self.chk_invert_nav.setChecked(
            bool(getattr(c.navigation, "invert_page_turn", True))
        )
        self.spin_click_x.setValue(c.navigation.click_x_offset)
        self.spin_click_y.setValue(c.navigation.click_y_offset)
        self.chk_refocus.setChecked(c.navigation.refocus_before_action)

        self.edit_hk_start.setText(c.hotkeys.start_stop)
        self.edit_hk_pause.setText(c.hotkeys.pause)
        self.edit_hk_emerg.setText(c.hotkeys.emergency_stop)

        self.edit_subfolder.setText(c.storage.book_subfolder_template)
        self.edit_filetpl.setText(c.storage.file_template)
        self.chk_resume.setChecked(c.storage.resume_from_existing)
        self.chk_auto_pdf.setChecked(bool(getattr(c.storage, "auto_pdf", True)))
        self.edit_pdf_name.setText(getattr(c.storage, "pdf_filename", "book.pdf"))

    def _on_accept(self) -> None:
        self._mgr.update(
            capture={
                "page_delay_sec": self.spin_delay.value(),
                "pre_capture_delay_sec": self.spin_pre_delay.value(),
                "save_format": self.cmb_format.currentText(),
                "jpeg_quality": self.spin_jpeg_q.value(),
                "webp_quality": self.spin_webp_q.value(),
                "skip_duplicates": self.chk_skip_dup.isChecked(),
                "duplicate_threshold": self.spin_dup_thresh.value(),
                "last_page_max_hash_distance": self.spin_last_page_dist.value(),
                "auto_stop_on_last_page": self.chk_auto_stop.isChecked(),
                "last_page_consecutive_dupes": self.spin_consec_dupes.value(),
                "min_captures_before_last_page_stop": self.spin_min_cap_last.value(),
                "max_duplicate_skips_without_save": self.spin_max_dup_skips.value(),
                "use_printwindow": self.chk_printwindow.isChecked(),
            },
            navigation={
                "method": self.cmb_nav_method.currentText(),
                "key": self.edit_nav_key.text().strip() or "pagedown",
                "key_delivery": self.cmb_key_delivery.currentText(),
                "wheel_scroll_notches": self.spin_wheel_notches.value(),
                "wheel_scroll_sign": int(self.cmb_wheel_sign.currentData()),
                "invert_page_turn": self.chk_invert_nav.isChecked(),
                "click_x_offset": self.spin_click_x.value(),
                "click_y_offset": self.spin_click_y.value(),
                "refocus_before_action": self.chk_refocus.isChecked(),
            },
            hotkeys={
                "start_stop": self.edit_hk_start.text().strip() or "f9",
                "pause": self.edit_hk_pause.text().strip() or "f10",
                "emergency_stop": self.edit_hk_emerg.text().strip() or "esc",
            },
            storage={
                "book_subfolder_template": self.edit_subfolder.text().strip() or "book_{timestamp}",
                "file_template": self.edit_filetpl.text().strip() or "page_{index:04d}",
                "resume_from_existing": self.chk_resume.isChecked(),
                "auto_pdf": self.chk_auto_pdf.isChecked(),
                "pdf_filename": self.edit_pdf_name.text().strip() or "book.pdf",
            },
        )
        self.accept()
