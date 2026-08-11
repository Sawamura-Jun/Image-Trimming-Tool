import datetime
import os
import sys

from PIL import Image, ImageGrab
from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtGui import (
    QColor,
    QCursor,
    QDoubleValidator,
    QFont,
    QIcon,
    QImage,
    QIntValidator,
    QKeySequence,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
    QShortcut,
)
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)


# アプリケーションウィンドウの主要GUIパラメータ
APP_TITLE = "Image Cropper"
APP_WINDOW_SIZE = (1120, 680)
APP_MINIMUM_SIZE = APP_WINDOW_SIZE
CONTROL_PANEL_WIDTH = 300
WINDOW_RESIZE_STEP = 0.2
PANEL_MARGIN = 14
PANEL_SPACING = 8
CONTROL_HEIGHT = 34
BUTTON_HEIGHT = 36
HANDLE_SIZE = 10
MIN_CROP_SIZE = 4
GRID_LINES = 20

# 色と表示スタイル
BACKGROUND_COLOR = QColor(35, 39, 47)
SURFACE_COLOR = "#20242c"
CARD_COLOR = "#292e38"
INPUT_COLOR = "#1f232b"
TEXT_COLOR = "#f4f6f8"
MUTED_TEXT_COLOR = "#aeb6c2"
ACCENT_COLOR = QColor(66, 184, 255)
ACCENT_HEX = "#42b8ff"

# 画像編集の既定値
DEFAULT_ROTATION_ANGLE = 0.1
DEFAULT_CROP_ASPECT = "1:1"
DEFAULT_IMAGE_SIZE = 1024
DEFAULT_JPEG_QUALITY = 70
CLIPBOARD_SAVE_DIR = r""  # 空ならPictures\Image-Cropperを使用する


APP_STYLE_SHEET = f"""
QMainWindow, QWidget#rootWidget {{
    background: {SURFACE_COLOR};
    color: {TEXT_COLOR};
    font-family: "Yu Gothic UI", "Meiryo UI", sans-serif;
    font-size: 14px;
}}
QFrame#controlPanel {{
    background: {SURFACE_COLOR};
    border-left: 1px solid #353b47;
}}
QFrame.sectionCard {{
    background: {CARD_COLOR};
    border: 1px solid #373e4b;
    border-radius: 12px;
}}
QLabel.sectionTitle {{
    color: {TEXT_COLOR};
    font-size: 15px;
    font-weight: 700;
}}
QLabel.fieldLabel {{
    color: {MUTED_TEXT_COLOR};
    font-size: 12px;
}}
QLineEdit {{
    min-height: {CONTROL_HEIGHT}px;
    padding: 0 10px;
    color: {TEXT_COLOR};
    background: {INPUT_COLOR};
    border: 1px solid #434b59;
    border-radius: 8px;
    selection-background-color: {ACCENT_HEX};
}}
QLineEdit:focus {{
    border: 1px solid {ACCENT_HEX};
}}
QPushButton {{
    min-height: {BUTTON_HEIGHT}px;
    padding: 0 14px;
    color: {TEXT_COLOR};
    background: #353c48;
    border: 1px solid #454e5d;
    border-radius: 9px;
    font-weight: 600;
}}
QPushButton:hover {{
    background: #414a58;
    border-color: #596576;
}}
QPushButton:pressed {{
    background: #2c323c;
}}
QPushButton#primaryButton {{
    color: #07131b;
    background: {ACCENT_HEX};
    border-color: {ACCENT_HEX};
}}
QPushButton#primaryButton:hover {{
    background: #69c7ff;
}}
QCheckBox {{
    color: {TEXT_COLOR};
    spacing: 8px;
}}
QCheckBox::indicator {{
    width: 18px;
    height: 18px;
}}
QCheckBox::indicator:unchecked {{
    background: {INPUT_COLOR};
    border: 1px solid #596272;
    border-radius: 5px;
}}
QCheckBox::indicator:checked {{
    background: {ACCENT_HEX};
    border: 1px solid {ACCENT_HEX};
    border-radius: 5px;
}}
QStatusBar {{
    color: {MUTED_TEXT_COLOR};
    background: #1b1f26;
    border-top: 1px solid #303641;
}}
"""


def resolve_clipboard_save_dir():
    """クリップボード画像の保存先ディレクトリを返す。"""
    if CLIPBOARD_SAVE_DIR:
        return CLIPBOARD_SAVE_DIR
    home = os.path.expanduser("~")
    if home and home != "~":
        return os.path.join(home, "Pictures", "Image-Cropper")
    # ホームディレクトリが解決できない場合だけカレントを使用する
    return os.path.join(os.getcwd(), "Image-Cropper")


def pil_to_qimage(pil_image):
    """Pillow画像を、元バッファに依存しないQImageへ変換する。"""
    rgba_image = pil_image.convert("RGBA")
    raw_data = rgba_image.tobytes("raw", "RGBA")
    image = QImage(
        raw_data,
        rgba_image.width,
        rgba_image.height,
        rgba_image.width * 4,
        QImage.Format.Format_RGBA8888,
    )
    return image.copy()


def qimage_to_pil(qimage):
    """クリップボードのQImageをPillow画像へ変換する。"""
    converted = qimage.convertToFormat(QImage.Format.Format_RGBA8888)
    width = converted.width()
    height = converted.height()
    byte_count = converted.bytesPerLine() * height
    data = bytes(converted.constBits()[:byte_count])
    return Image.frombuffer(
        "RGBA",
        (width, height),
        data,
        "raw",
        "RGBA",
        converted.bytesPerLine(),
        1,
    ).copy()


def parse_aspect_ratio(value):
    """「幅:高さ」を検証し、比率を返す。"""
    width_ratio, height_ratio = map(float, value.split(":"))
    if width_ratio <= 0 or height_ratio <= 0:
        raise ValueError("縦横比には正の値が必要です")
    return width_ratio / height_ratio


class ImagePanel(QWidget):
    """画像、グリッド、トリミング範囲を描画するキャンバス。"""

    HANDLE_SIZE = HANDLE_SIZE
    MIN_CROP_SIZE = MIN_CROP_SIZE

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setMinimumSize(420, 360)

        # 表示用キャッシュ
        self._cached_pixmap = None
        self._cached_size = (0, 0)
        self._cached_image_id = None

        # 画像と編集履歴
        self.original_image = None
        self.current_image = None
        self.rotation_base_image = None
        self.rotation_angle_total = 0.0
        self.file_name = ""
        self.file_dir = ""
        self.from_clipboard = False
        self.crop_history = []
        self.max_crop_history = 10

        # トリミング範囲の操作状態
        self.crop_rect = None
        self.crop_aspect = DEFAULT_CROP_ASPECT
        self.fixed_aspect = True
        self.mode = "idle"
        self.drag_handle = None
        self.drag_start = QPoint()
        self.original_rect = None

        # 画像の表示位置と表示寸法
        self.display_offset_x = 0
        self.display_offset_y = 0
        self.display_width = 0
        self.display_height = 0
        self.old_display_width = 0
        self.old_display_height = 0

    def resizeEvent(self, event):
        """キャンバスのリサイズに追従して選択範囲も拡縮する。"""
        self.old_display_width = self.display_width
        self.old_display_height = self.display_height
        self.UpdateDisplayGeometry()
        self.RescaleCropRect()
        self.update()
        super().resizeEvent(event)

    def RescaleCropRect(self):
        if (
            not self.crop_rect
            or self.old_display_width == 0
            or self.old_display_height == 0
            or self.display_width == 0
            or self.display_height == 0
        ):
            return
        old_x, old_y, old_w, old_h = self.crop_rect
        scale_x = self.display_width / self.old_display_width
        scale_y = self.display_height / self.old_display_height
        new_x = old_x * scale_x
        new_y = old_y * scale_y
        new_w = old_w * scale_x
        new_h = old_h * scale_y
        rect = self._ensure_min_size(
            QRect(
                round(new_x),
                round(new_y),
                round(new_w),
                round(new_h),
            )
        )
        self.crop_rect = (rect.x(), rect.y(), rect.width(), rect.height())

    def ClipRect(self, x, y, width, height):
        """矩形を画像の表示エリア内に収める。"""
        width = max(0, min(width, self.display_width))
        height = max(0, min(height, self.display_height))
        x = max(0, min(x, self.display_width - width))
        y = max(0, min(y, self.display_height - height))
        return x, y, width, height

    def UpdateDisplayGeometry(self):
        if self.current_image is None:
            self.display_offset_x = 0
            self.display_offset_y = 0
            self.display_width = 0
            self.display_height = 0
            return
        panel_width = max(1, self.width())
        panel_height = max(1, self.height())
        image_width, image_height = self.current_image.size
        scale = min(panel_width / image_width, panel_height / image_height)
        self.display_width = max(1, round(image_width * scale))
        self.display_height = max(1, round(image_height * scale))
        self.display_offset_x = (panel_width - self.display_width) // 2
        self.display_offset_y = (panel_height - self.display_height) // 2
        self._cached_pixmap = None

    def _event_to_display_point(self, event):
        point = event.position().toPoint()
        return QPoint(
            point.x() - self.display_offset_x,
            point.y() - self.display_offset_y,
        )

    def _clamp_display_point(self, point):
        return QPoint(
            max(0, min(point.x(), self.display_width)),
            max(0, min(point.y(), self.display_height)),
        )

    def _point_in_display(self, point):
        return (
            0 <= point.x() <= self.display_width
            and 0 <= point.y() <= self.display_height
        )

    @staticmethod
    def _rect_contains_point(rect_tuple, point):
        x, y, width, height = rect_tuple
        return x <= point.x() <= x + width and y <= point.y() <= y + height

    def _iter_handle_rects_display(self):
        if not self.crop_rect:
            return []
        half = self.HANDLE_SIZE // 2
        x, y, width, height = self.crop_rect
        points = {
            "top_left": (x, y),
            "top": (x + width / 2, y),
            "top_right": (x + width, y),
            "right": (x + width, y + height / 2),
            "bottom_right": (x + width, y + height),
            "bottom": (x + width / 2, y + height),
            "bottom_left": (x, y + height),
            "left": (x, y + height / 2),
        }
        return [
            (
                name,
                QRect(
                    round(point_x) - half,
                    round(point_y) - half,
                    self.HANDLE_SIZE,
                    self.HANDLE_SIZE,
                ),
            )
            for name, (point_x, point_y) in points.items()
        ]

    def _iter_handle_rects_panel(self):
        for name, rect in self._iter_handle_rects_display():
            translated = QRect(rect)
            translated.translate(self.display_offset_x, self.display_offset_y)
            yield name, translated

    def _hit_test_handle(self, point):
        for name, rect in self._iter_handle_rects_display():
            if rect.contains(point):
                return name
        return None

    def _get_aspect_ratio(self):
        if not self.fixed_aspect:
            return None
        try:
            return parse_aspect_ratio(self.crop_aspect)
        except (TypeError, ValueError):
            return None

    def _ensure_within_display(self, rect):
        if rect is None or self.display_width <= 0 or self.display_height <= 0:
            return QRect()
        bounded = QRect(rect)
        width = min(max(1, bounded.width()), self.display_width)
        height = min(max(1, bounded.height()), self.display_height)
        x = max(0, min(bounded.x(), self.display_width - width))
        y = max(0, min(bounded.y(), self.display_height - height))
        return QRect(x, y, width, height)

    def _ensure_min_size(self, rect):
        if self.display_width <= 0 or self.display_height <= 0:
            return QRect()
        bounded = self._ensure_within_display(rect)
        width = min(max(self.MIN_CROP_SIZE, bounded.width()), self.display_width)
        height = min(max(self.MIN_CROP_SIZE, bounded.height()), self.display_height)
        x = max(0, min(bounded.x(), self.display_width - width))
        y = max(0, min(bounded.y(), self.display_height - height))
        return QRect(x, y, width, height)

    def _create_rect_with_ratio(self, anchor, current, ratio):
        delta_x = current.x() - anchor.x()
        delta_y = current.y() - anchor.y()
        absolute_x = abs(delta_x)
        absolute_y = abs(delta_y)
        if absolute_x == 0 and absolute_y == 0:
            return QRect(anchor.x(), anchor.y(), 0, 0)
        if absolute_y == 0:
            absolute_y = round(absolute_x / ratio)
        if absolute_x == 0:
            absolute_x = round(absolute_y * ratio)
        if absolute_y and absolute_x / absolute_y > ratio:
            absolute_x = round(absolute_y * ratio)
        else:
            absolute_y = round(absolute_x / ratio)
        target_x = anchor.x() + (absolute_x if delta_x >= 0 else -absolute_x)
        target_y = anchor.y() + (absolute_y if delta_y >= 0 else -absolute_y)
        return self._ensure_within_display(
            QRect(
                min(anchor.x(), target_x),
                min(anchor.y(), target_y),
                abs(target_x - anchor.x()),
                abs(target_y - anchor.y()),
            )
        )

    def _create_rect(self, anchor, current):
        anchor = self._clamp_display_point(anchor)
        current = self._clamp_display_point(current)
        ratio = self._get_aspect_ratio()
        if ratio:
            rect = self._create_rect_with_ratio(anchor, current, ratio)
        else:
            rect = QRect(
                min(anchor.x(), current.x()),
                min(anchor.y(), current.y()),
                abs(current.x() - anchor.x()),
                abs(current.y() - anchor.y()),
            )
        return self._ensure_min_size(rect)

    def _rect_from_crop(self):
        if not self.crop_rect:
            return None
        x, y, width, height = self.crop_rect
        return QRect(round(x), round(y), round(width), round(height))

    def _set_crop_rect(self, rect):
        rect = self._ensure_min_size(rect)
        self.crop_rect = (rect.x(), rect.y(), rect.width(), rect.height())

    def _update_selection_creation(self, anchor, current):
        self._set_crop_rect(self._create_rect(anchor, current))

    def _update_selection_move(self, delta_x, delta_y):
        if not self.original_rect:
            return
        rect = QRect(self.original_rect)
        rect.translate(delta_x, delta_y)
        self._set_crop_rect(rect)

    def _rect_from_horizontal_anchor(self, anchor_x, width, origin, to_left, ratio):
        width = max(self.MIN_CROP_SIZE, min(width, self.display_width))
        height = max(self.MIN_CROP_SIZE, round(width / ratio))
        if height > self.display_height:
            height = self.display_height
            width = max(self.MIN_CROP_SIZE, round(height * ratio))
        center_y = origin.y() + origin.height() / 2
        top = max(0, min(round(center_y - height / 2), self.display_height - height))
        if to_left:
            right = min(anchor_x, self.display_width)
            left = max(0, right - width)
        else:
            left = max(0, anchor_x)
            right = min(self.display_width, left + width)
            left = right - width
        return QRect(round(left), round(top), round(right - left), round(height))

    def _rect_from_vertical_anchor(self, anchor_y, height, origin, to_top, ratio):
        height = max(self.MIN_CROP_SIZE, min(height, self.display_height))
        width = max(self.MIN_CROP_SIZE, round(height * ratio))
        if width > self.display_width:
            width = self.display_width
            height = max(self.MIN_CROP_SIZE, round(width / ratio))
        center_x = origin.x() + origin.width() / 2
        left = max(0, min(round(center_x - width / 2), self.display_width - width))
        if to_top:
            bottom = min(anchor_y, self.display_height)
            top = max(0, bottom - height)
        else:
            top = max(0, anchor_y)
            bottom = min(self.display_height, top + height)
            top = bottom - height
        return QRect(round(left), round(top), round(width), round(bottom - top))

    def _resize_corner_with_ratio(self, point, handle, origin, ratio):
        origin_right = origin.x() + origin.width()
        origin_bottom = origin.y() + origin.height()
        if handle == "top_left":
            anchor = QPoint(origin_right, origin_bottom)
            horizontal, vertical = -1, -1
        elif handle == "top_right":
            anchor = QPoint(origin.x(), origin_bottom)
            horizontal, vertical = 1, -1
        elif handle == "bottom_left":
            anchor = QPoint(origin_right, origin.y())
            horizontal, vertical = -1, 1
        else:
            anchor = QPoint(origin.x(), origin.y())
            horizontal, vertical = 1, 1
        width = max(self.MIN_CROP_SIZE, abs(point.x() - anchor.x()))
        height = round(width / ratio)
        available_height = abs(point.y() - anchor.y())
        if available_height and height > available_height:
            height = max(self.MIN_CROP_SIZE, available_height)
            width = round(height * ratio)
        left = anchor.x() - width if horizontal < 0 else anchor.x()
        top = anchor.y() - height if vertical < 0 else anchor.y()
        return self._ensure_min_size(QRect(left, top, width, height))

    def _resize_with_ratio(self, point, handle, ratio):
        if not self.original_rect:
            return QRect()
        origin = QRect(self.original_rect)
        origin_right = origin.x() + origin.width()
        origin_bottom = origin.y() + origin.height()
        if handle == "left":
            width = max(self.MIN_CROP_SIZE, min(origin_right - point.x(), origin_right))
            rect = self._rect_from_horizontal_anchor(origin_right, width, origin, True, ratio)
        elif handle == "right":
            width = max(
                self.MIN_CROP_SIZE,
                min(point.x() - origin.x(), self.display_width - origin.x()),
            )
            rect = self._rect_from_horizontal_anchor(origin.x(), width, origin, False, ratio)
        elif handle == "top":
            height = max(self.MIN_CROP_SIZE, min(origin_bottom - point.y(), origin_bottom))
            rect = self._rect_from_vertical_anchor(origin_bottom, height, origin, True, ratio)
        elif handle == "bottom":
            height = max(
                self.MIN_CROP_SIZE,
                min(point.y() - origin.y(), self.display_height - origin.y()),
            )
            rect = self._rect_from_vertical_anchor(origin.y(), height, origin, False, ratio)
        else:
            rect = self._resize_corner_with_ratio(point, handle, origin, ratio)
        return self._ensure_min_size(rect)

    def _resize_free(self, point, handle):
        rect = QRect(self.original_rect)
        left = rect.x()
        top = rect.y()
        right = rect.x() + rect.width()
        bottom = rect.y() + rect.height()
        if "left" in handle:
            left = min(point.x(), right - self.MIN_CROP_SIZE)
        if "right" in handle:
            right = max(point.x(), left + self.MIN_CROP_SIZE)
        if "top" in handle:
            top = min(point.y(), bottom - self.MIN_CROP_SIZE)
        if "bottom" in handle:
            bottom = max(point.y(), top + self.MIN_CROP_SIZE)
        left = max(0, min(left, self.display_width))
        right = max(0, min(right, self.display_width))
        top = max(0, min(top, self.display_height))
        bottom = max(0, min(bottom, self.display_height))
        return self._ensure_min_size(QRect(left, top, right - left, bottom - top))

    def _update_selection_resize(self, point):
        if not self.original_rect or not self.drag_handle:
            return
        ratio = self._get_aspect_ratio() if self.fixed_aspect else None
        if ratio:
            rect = self._resize_with_ratio(point, self.drag_handle, ratio)
        else:
            rect = self._resize_free(point, self.drag_handle)
        self._set_crop_rect(rect)

    def _update_cursor(self, display_point):
        handle = self._hit_test_handle(display_point)
        if handle in ("top_left", "bottom_right"):
            shape = Qt.CursorShape.SizeFDiagCursor
        elif handle in ("top_right", "bottom_left"):
            shape = Qt.CursorShape.SizeBDiagCursor
        elif handle in ("left", "right"):
            shape = Qt.CursorShape.SizeHorCursor
        elif handle in ("top", "bottom"):
            shape = Qt.CursorShape.SizeVerCursor
        elif self._point_in_display(display_point):
            if self.crop_rect and self._rect_contains_point(self.crop_rect, display_point):
                shape = Qt.CursorShape.SizeAllCursor
            else:
                shape = Qt.CursorShape.CrossCursor
        else:
            shape = Qt.CursorShape.ArrowCursor
        self.setCursor(QCursor(shape))

    def ApplyAspectRatioToSelection(self):
        if not self.fixed_aspect or not self.crop_rect:
            return
        ratio = self._get_aspect_ratio()
        if not ratio:
            return
        x, y, width, height = self.crop_rect
        center_x = x + width / 2
        center_y = y + height / 2
        target_height = round(width / ratio)
        if target_height <= self.display_height:
            height = target_height
        width = round(height * ratio)
        if width > self.display_width:
            width = self.display_width
            height = round(width / ratio)
        if height > self.display_height:
            height = self.display_height
            width = round(height * ratio)
        rect = QRect(
            round(center_x - width / 2),
            round(center_y - height / 2),
            max(self.MIN_CROP_SIZE, width),
            max(self.MIN_CROP_SIZE, height),
        )
        self._set_crop_rect(rect)

    def SetImage(self, pil_image, file_name=""):
        """新しい画像を設定し、編集状態を初期化する。"""
        if pil_image is None:
            return
        self.from_clipboard = False
        self.original_image = pil_image.copy()
        self.current_image = pil_image.copy()
        self.crop_history = [self.current_image.copy()]
        self.rotation_base_image = self.current_image.copy()
        self.rotation_angle_total = 0.0
        self.mode = "idle"
        self.drag_handle = None
        self.original_rect = None
        self.drag_start = QPoint()
        self.file_name = os.path.basename(file_name)
        self.file_dir = os.path.dirname(file_name)
        self.UpdateDisplayGeometry()
        self.InitCropRect()
        self.UpdateTitle()
        self._cached_pixmap = None
        self.update()

    def UpdateTitle(self):
        top_window = self.window()
        if self.file_name and self.current_image is not None:
            width, height = self.current_image.size
            top_window.setWindowTitle(f"{self.file_name}  —  {width} × {height}")
        elif top_window:
            top_window.setWindowTitle(APP_TITLE)

    def _update_pixmap_cache(self):
        cache_key = (self.display_width, self.display_height)
        if (
            self._cached_pixmap is None
            or self._cached_size != cache_key
            or self._cached_image_id != id(self.current_image)
        ):
            qimage = pil_to_qimage(self.current_image)
            self._cached_pixmap = QPixmap.fromImage(qimage).scaled(
                self.display_width,
                self.display_height,
                Qt.AspectRatioMode.IgnoreAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            self._cached_size = cache_key
            self._cached_image_id = id(self.current_image)

    def _paint_empty_state(self, painter):
        """画像未読込時に操作案内を中央表示する。"""
        center_x = self.width() // 2
        center_y = self.height() // 2
        drop_rect = QRect(center_x - 230, center_y - 110, 460, 220)
        pen = QPen(QColor(82, 91, 106), 2, Qt.PenStyle.DashLine)
        painter.setPen(pen)
        painter.setBrush(QColor(41, 46, 56))
        painter.drawRoundedRect(drop_rect, 18, 18)
        painter.setPen(QColor(TEXT_COLOR))
        title_font = QFont("Yu Gothic UI", 17, QFont.Weight.DemiBold)
        painter.setFont(title_font)
        painter.drawText(
            drop_rect.adjusted(20, 52, -20, -90),
            Qt.AlignmentFlag.AlignCenter,
            "画像をここにドロップ",
        )
        painter.setPen(QColor(MUTED_TEXT_COLOR))
        painter.setFont(QFont("Yu Gothic UI", 10))
        painter.drawText(
            drop_rect.adjusted(20, 105, -20, -30),
            Qt.AlignmentFlag.AlignCenter,
            "または Ctrl + V でクリップボードから貼り付け",
        )

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.fillRect(self.rect(), BACKGROUND_COLOR)
        if self.current_image is None:
            self._paint_empty_state(painter)
            return

        self._update_pixmap_cache()
        image_rect = QRect(
            self.display_offset_x,
            self.display_offset_y,
            self.display_width,
            self.display_height,
        )
        painter.drawPixmap(image_rect, self._cached_pixmap)

        # 構図を確認しやすい薄いグリッドを描画する
        grid_pen = QPen(QColor(255, 255, 255, 66), 1, Qt.PenStyle.DotLine)
        painter.setPen(grid_pen)
        for index in range(GRID_LINES + 1):
            horizontal_y = self.display_offset_y + round(self.display_height * index / GRID_LINES)
            vertical_x = self.display_offset_x + round(self.display_width * index / GRID_LINES)
            painter.drawLine(
                self.display_offset_x,
                horizontal_y,
                self.display_offset_x + self.display_width,
                horizontal_y,
            )
            painter.drawLine(
                vertical_x,
                self.display_offset_y,
                vertical_x,
                self.display_offset_y + self.display_height,
            )

        if not self.crop_rect:
            return
        crop_x, crop_y, crop_width, crop_height = self.crop_rect
        crop_panel_rect = QRect(
            round(crop_x + self.display_offset_x),
            round(crop_y + self.display_offset_y),
            round(crop_width),
            round(crop_height),
        )

        # 選択範囲の外側だけを半透明で暗くする
        overlay = QPainterPath()
        overlay.setFillRule(Qt.FillRule.OddEvenFill)
        overlay.addRect(self.rect())
        overlay.addRect(crop_panel_rect)
        painter.fillPath(overlay, QColor(0, 0, 0, 145))
        painter.setPen(QPen(ACCENT_COLOR, 2))
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRect(crop_panel_rect)

        # リサイズ用ハンドルは背景から見分けやすい白で描く
        painter.setPen(QPen(ACCENT_COLOR, 1))
        painter.setBrush(QColor(248, 250, 252))
        for _, handle_rect in self._iter_handle_rects_panel():
            painter.drawRoundedRect(handle_rect, 2, 2)

    def mousePressEvent(self, event):
        if self.current_image is None or event.button() != Qt.MouseButton.LeftButton:
            return
        display_point = self._event_to_display_point(event)
        handle = self._hit_test_handle(display_point)
        if handle and self.crop_rect:
            self.mode = "resizing"
            self.drag_handle = handle
            self.original_rect = self._rect_from_crop()
        elif self.crop_rect and self._rect_contains_point(self.crop_rect, display_point):
            self.mode = "moving"
            self.drag_handle = "inside"
            self.original_rect = self._rect_from_crop()
        elif self._point_in_display(display_point):
            self.mode = "creating"
            self.drag_handle = None
            self.original_rect = None
            anchor = self._clamp_display_point(display_point)
            self.crop_rect = (anchor.x(), anchor.y(), 0, 0)
        else:
            return
        self.drag_start = self._clamp_display_point(display_point)
        self.grabMouse()
        self.update()

    def mouseMoveEvent(self, event):
        if self.current_image is None:
            return
        display_point = self._event_to_display_point(event)
        if event.buttons() & Qt.MouseButton.LeftButton and self.mode != "idle":
            point = self._clamp_display_point(display_point)
            if self.mode == "creating":
                self._update_selection_creation(self.drag_start, point)
            elif self.mode == "moving" and self.original_rect:
                self._update_selection_move(
                    point.x() - self.drag_start.x(),
                    point.y() - self.drag_start.y(),
                )
            elif self.mode == "resizing" and self.original_rect:
                self._update_selection_resize(point)
            self.update()
            return
        self._update_cursor(display_point)

    def mouseReleaseEvent(self, event):
        if event.button() != Qt.MouseButton.LeftButton:
            return
        previous_mode = self.mode
        if QWidget.mouseGrabber() is self:
            self.releaseMouse()
        self.mode = "idle"
        self.drag_handle = None
        self.original_rect = None
        self.drag_start = QPoint()
        if previous_mode == "creating" and self.crop_rect:
            self._set_crop_rect(self._rect_from_crop())
        self._update_cursor(self._event_to_display_point(event))
        self.update()

    def leaveEvent(self, event):
        if self.mode == "idle":
            self.setCursor(QCursor(Qt.CursorShape.ArrowCursor))
        super().leaveEvent(event)

    def wheelEvent(self, event):
        top_window = self.window()
        if hasattr(top_window, "OnMouseWheelResize"):
            top_window.OnMouseWheelResize(event)
        else:
            super().wheelEvent(event)

    def RotateImage(self, delta):
        if self.rotation_base_image is None:
            return
        self.rotation_angle_total = (self.rotation_angle_total + delta) % 360
        self.current_image = self.rotation_base_image.rotate(
            self.rotation_angle_total,
            expand=True,
            resample=Image.Resampling.BICUBIC,
        )
        self._append_history()
        self.UpdateDisplayGeometry()
        self.InitCropRect()
        self.UpdateTitle()
        self.update()

    def _append_history(self):
        if len(self.crop_history) >= self.max_crop_history:
            self.crop_history.pop(0)
        self.crop_history.append(self.current_image.copy())

    def CropImage(self):
        if not self.crop_rect or self.current_image is None:
            return
        if self.display_width == 0 or self.display_height == 0:
            return
        image_width, image_height = self.current_image.size
        crop_x, crop_y, rect_width, rect_height = self.crop_rect
        scale_x = image_width / self.display_width
        scale_y = image_height / self.display_height
        left = max(0, round(crop_x * scale_x))
        top = max(0, round(crop_y * scale_y))
        right = min(image_width, round((crop_x + rect_width) * scale_x))
        bottom = min(image_height, round((crop_y + rect_height) * scale_y))
        if right <= left or bottom <= top:
            return
        self.current_image = self.current_image.crop((left, top, right, bottom))
        self._append_history()
        self.rotation_base_image = self.current_image.copy()
        self.rotation_angle_total = 0.0
        self.UpdateDisplayGeometry()
        self.InitCropRect()
        self.UpdateTitle()
        self.update()

    def RevertCrop(self):
        if len(self.crop_history) <= 1:
            return
        self.crop_history.pop()
        self.current_image = self.crop_history[-1].copy()
        self.rotation_base_image = self.current_image.copy()
        self.rotation_angle_total = 0.0
        self.UpdateDisplayGeometry()
        self.InitCropRect()
        self.UpdateTitle()
        self.update()

    def ResizeImage(self, target_size):
        if target_size <= 0:
            raise ValueError("画像サイズには正の値が必要です")
        if self.current_image is None:
            return
        width, height = self.current_image.size
        long_side = max(width, height)
        if long_side <= target_size:
            return
        ratio = target_size / long_side
        new_size = (max(1, round(width * ratio)), max(1, round(height * ratio)))
        self.current_image = self.current_image.resize(new_size, Image.Resampling.LANCZOS)
        self._append_history()
        self.rotation_base_image = self.current_image.copy()
        self.rotation_angle_total = 0.0
        self.UpdateDisplayGeometry()
        self.InitCropRect()
        self.UpdateTitle()
        self.update()

    def SaveImage(self, jpeg_quality):
        if self.current_image is None:
            return None
        if not 1 <= jpeg_quality <= 100:
            raise ValueError("JPG品質は1から100で指定してください")
        if self.from_clipboard:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            save_dir = resolve_clipboard_save_dir()
            os.makedirs(save_dir, exist_ok=True)
            save_path = os.path.join(save_dir, f"clipboard_{timestamp}.png")
            self.current_image.save(save_path, "PNG")
            return save_path
        if not self.file_name:
            return None
        name, extension = os.path.splitext(self.file_name)
        save_path = os.path.join(self.file_dir, f"{name}_trm{extension}")
        image_to_save = self.current_image
        save_parameters = {}
        if extension.lower() in (".jpg", ".jpeg"):
            image_to_save = image_to_save.convert("RGB")
            save_parameters["quality"] = jpeg_quality
        image_to_save.save(save_path, **save_parameters)
        return save_path

    def InitCropRect(self):
        if self.display_width <= 0 or self.display_height <= 0:
            self.crop_rect = None
            return
        try:
            ratio = self._get_aspect_ratio() if self.fixed_aspect else None
            if ratio:
                if ratio < 1:
                    rect_height = max(self.MIN_CROP_SIZE, self.display_height // 4)
                    rect_width = round(rect_height * ratio)
                else:
                    rect_width = max(self.MIN_CROP_SIZE, self.display_width // 4)
                    rect_height = round(rect_width / ratio)
            else:
                rect_width = max(self.MIN_CROP_SIZE, self.display_width // 4)
                rect_height = min(rect_width, self.display_height)
        except (TypeError, ValueError):
            rect_width = max(self.MIN_CROP_SIZE, self.display_width // 4)
            rect_height = min(rect_width, self.display_height)
        rect = QRect(
            round((self.display_width - rect_width) / 2),
            round((self.display_height - rect_height) / 2),
            rect_width,
            rect_height,
        )
        self._set_crop_rect(rect)
        self.mode = "idle"
        self.drag_handle = None
        self.original_rect = None
        self.drag_start = QPoint()


class ControlPanel(QFrame):
    """画像編集操作をまとめた右側のコントロールパネル。"""

    def __init__(self, parent, image_panel):
        super().__init__(parent)
        self.image_panel = image_panel
        self.setObjectName("controlPanel")
        self.setFixedWidth(CONTROL_PANEL_WIDTH)
        self.InitUI()

    @staticmethod
    def _section(title):
        card = QFrame()
        card.setProperty("class", "sectionCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(6)
        title_label = QLabel(title)
        title_label.setProperty("class", "sectionTitle")
        layout.addWidget(title_label)
        return card, layout

    @staticmethod
    def _field_row(label_text, editor):
        layout = QHBoxLayout()
        layout.setSpacing(10)
        label = QLabel(label_text)
        label.setProperty("class", "fieldLabel")
        label.setMinimumWidth(80)
        layout.addWidget(label)
        editor.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(editor, 1)
        return layout

    @staticmethod
    def _button(text, primary=False):
        button = QPushButton(text)
        if primary:
            button.setObjectName("primaryButton")
        button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        return button

    def InitUI(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(PANEL_MARGIN, PANEL_MARGIN, PANEL_MARGIN, PANEL_MARGIN)
        root_layout.setSpacing(PANEL_SPACING)

        header = QLabel("編集ツール")
        header.setFont(QFont("Yu Gothic UI", 18, QFont.Weight.Bold))
        root_layout.addWidget(header)

        # 回転セクション
        rotate_card, rotate_layout = self._section("回転")
        self.tc_rot = QLineEdit(str(DEFAULT_ROTATION_ANGLE))
        self.tc_rot.setValidator(QDoubleValidator(-360.0, 360.0, 3, self))
        self.tc_rot.setToolTip("1回のクリックで回転する角度")
        rotate_layout.addLayout(self._field_row("角度 (°)", self.tc_rot))
        rotate_buttons = QHBoxLayout()
        rotate_buttons.setSpacing(8)
        left_button = self._button("↶  左回転")
        right_button = self._button("右回転  ↷")
        left_button.clicked.connect(self.OnRotateLeft)
        right_button.clicked.connect(self.OnRotateRight)
        rotate_buttons.addWidget(left_button)
        rotate_buttons.addWidget(right_button)
        rotate_layout.addLayout(rotate_buttons)
        root_layout.addWidget(rotate_card)

        # トリミングセクション
        crop_card, crop_layout = self._section("トリミング")
        self.cb_aspect = QCheckBox("縦横比を固定")
        self.cb_aspect.setChecked(True)
        self.cb_aspect.toggled.connect(self.OnAspectCheckbox)
        crop_layout.addWidget(self.cb_aspect)
        self.tc_crop = QLineEdit(DEFAULT_CROP_ASPECT)
        self.tc_crop.setToolTip("例: 1:1、16:9、4:3")
        self.tc_crop.returnPressed.connect(self.OnAspectEnter)
        crop_layout.addLayout(self._field_row("縦横比", self.tc_crop))
        crop_button = self._button("選択範囲でトリミング", primary=True)
        crop_button.clicked.connect(self.OnCrop)
        revert_button = self._button("ひとつ前に戻す")
        revert_button.clicked.connect(self.OnRevert)
        crop_layout.addWidget(crop_button)
        crop_layout.addWidget(revert_button)
        root_layout.addWidget(crop_card)

        # 出力セクション
        output_card, output_layout = self._section("サイズ・保存")
        self.tc_resize = QLineEdit(str(DEFAULT_IMAGE_SIZE))
        self.tc_resize.setValidator(QIntValidator(1, 100000, self))
        output_layout.addLayout(self._field_row("長辺 (px)", self.tc_resize))
        resize_button = self._button("画像サイズを変更")
        resize_button.clicked.connect(self.OnResizeImage)
        output_layout.addWidget(resize_button)
        self.tc_quality = QLineEdit(str(DEFAULT_JPEG_QUALITY))
        self.tc_quality.setValidator(QIntValidator(1, 100, self))
        output_layout.addLayout(self._field_row("JPG品質", self.tc_quality))
        save_button = self._button("画像を保存", primary=True)
        save_button.clicked.connect(self.OnSave)
        output_layout.addWidget(save_button)
        root_layout.addWidget(output_card)

        root_layout.addItem(
            QSpacerItem(0, 0, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        )
        shortcut_hint = QLabel("Ctrl + V  貼り付け    Ctrl + C  コピー")
        shortcut_hint.setProperty("class", "fieldLabel")
        shortcut_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root_layout.addWidget(shortcut_hint)

    def _show_error(self, message):
        QMessageBox.critical(self, "入力エラー", message)

    def _show_status(self, message):
        top_window = self.window()
        if isinstance(top_window, QMainWindow):
            top_window.statusBar().showMessage(message, 5000)

    def OnRotateLeft(self, _checked=False):
        try:
            self.image_panel.RotateImage(float(self.tc_rot.text()))
        except ValueError:
            self._show_error("回転角度に数値を入力してください。")

    def OnRotateRight(self, _checked=False):
        try:
            self.image_panel.RotateImage(-float(self.tc_rot.text()))
        except ValueError:
            self._show_error("回転角度に数値を入力してください。")

    def OnAspectEnter(self):
        try:
            parse_aspect_ratio(self.tc_crop.text())
            self.image_panel.crop_aspect = self.tc_crop.text()
            self.image_panel.fixed_aspect = self.cb_aspect.isChecked()
            if self.image_panel.crop_rect:
                self.image_panel.ApplyAspectRatioToSelection()
            else:
                self.image_panel.InitCropRect()
            self.image_panel.update()
        except (TypeError, ValueError):
            self._show_error("縦横比の入力形式が不正です。例: 1:1")

    def OnAspectCheckbox(self, checked):
        self.image_panel.fixed_aspect = checked
        self.tc_crop.setEnabled(checked)
        if checked:
            self.OnAspectEnter()
        else:
            self.image_panel.update()

    def OnCrop(self, _checked=False):
        if self.cb_aspect.isChecked():
            try:
                parse_aspect_ratio(self.tc_crop.text())
                self.image_panel.crop_aspect = self.tc_crop.text()
                self.image_panel.ApplyAspectRatioToSelection()
            except (TypeError, ValueError):
                self._show_error("縦横比の入力形式が不正です。例: 1:1")
                return
        self.image_panel.CropImage()
        self._show_status("トリミングを適用しました")

    def OnRevert(self, _checked=False):
        self.image_panel.RevertCrop()
        self._show_status("ひとつ前の状態に戻しました")

    def OnResizeImage(self, _checked=False):
        try:
            self.image_panel.ResizeImage(int(self.tc_resize.text()))
            self._show_status("画像サイズを変更しました")
        except ValueError:
            self._show_error("画像サイズに正の整数を入力してください。")

    def OnSave(self, _checked=False):
        try:
            save_path = self.image_panel.SaveImage(int(self.tc_quality.text()))
            if save_path:
                self._show_status(f"保存しました: {save_path}")
            else:
                self._show_status("保存する画像がありません")
        except (OSError, ValueError) as error:
            self._show_error(f"画像を保存できませんでした。\n{error}")

    def wheelEvent(self, event):
        top_window = self.window()
        if hasattr(top_window, "OnMouseWheelResize"):
            top_window.OnMouseWheelResize(event)
        else:
            super().wheelEvent(event)


class ImageEditorFrame(QMainWindow):
    """アプリ全体のメインウィンドウ。"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(*APP_WINDOW_SIZE)
        self.setMinimumSize(*APP_MINIMUM_SIZE)
        self.setAcceptDrops(True)
        icon_path = os.path.join(os.path.dirname(__file__), "ico", "ImageCrop.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.InitUI()
        self._create_shortcuts()
        self._center_on_screen()

    def InitUI(self):
        root_widget = QWidget()
        root_widget.setObjectName("rootWidget")
        root_layout = QHBoxLayout(root_widget)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)
        self.image_panel = ImagePanel(root_widget)
        self.control_panel = ControlPanel(root_widget, self.image_panel)
        root_layout.addWidget(self.image_panel, 1)
        root_layout.addWidget(self.control_panel)
        self.setCentralWidget(root_widget)

        status_bar = QStatusBar(self)
        status_bar.setSizeGripEnabled(False)
        status_bar.showMessage("画像をドロップするか、Ctrl + V で貼り付けてください")
        self.setStatusBar(status_bar)

    def _create_shortcuts(self):
        self.paste_shortcut = QShortcut(QKeySequence.StandardKey.Paste, self)
        self.copy_shortcut = QShortcut(QKeySequence.StandardKey.Copy, self)
        self.paste_shortcut.activated.connect(self.PasteImageFromClipboard)
        self.copy_shortcut.activated.connect(self.CopyImageToClipboard)

    def _center_on_screen(self):
        screen = QApplication.primaryScreen()
        if screen:
            area = screen.availableGeometry()
            geometry = self.frameGeometry()
            geometry.moveCenter(area.center())
            self.move(geometry.topLeft())

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and any(
            url.isLocalFile() for url in event.mimeData().urls()
        ):
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        if not paths:
            return
        try:
            with Image.open(paths[0]) as source_image:
                image = source_image.copy()
            self.image_panel.SetImage(image, file_name=paths[0])
            self.statusBar().showMessage(f"読み込みました: {paths[0]}", 5000)
            event.acceptProposedAction()
        except (OSError, ValueError) as error:
            QMessageBox.critical(self, "読込エラー", f"画像を読み込めませんでした。\n{error}")

    def _get_display_client_area(self):
        screen = self.screen() or QApplication.primaryScreen()
        return screen.availableGeometry() if screen else QRect(0, 0, *APP_WINDOW_SIZE)

    def _clamp_scale(self, target_scale):
        base_width, base_height = APP_WINDOW_SIZE
        display_area = self._get_display_client_area()
        max_scale = min(
            display_area.width() / base_width,
            display_area.height() / base_height,
        )
        return max(1.0, min(target_scale, max(1.0, max_scale)))

    def _resize_and_center(self, scale):
        base_width, base_height = APP_WINDOW_SIZE
        new_width = round(base_width * scale)
        new_height = round(base_height * scale)
        display_area = self._get_display_client_area()
        new_x = display_area.x() + (display_area.width() - new_width) // 2
        new_y = display_area.y() + (display_area.height() - new_height) // 2
        self.setGeometry(new_x, new_y, new_width, new_height)

    def OnMouseWheelResize(self, event):
        rotation = event.angleDelta().y()
        if rotation == 0:
            event.ignore()
            return
        steps = rotation / 120
        current_scale = self.width() / APP_WINDOW_SIZE[0]
        target_scale = current_scale + WINDOW_RESIZE_STEP * steps
        self._resize_and_center(self._clamp_scale(target_scale))
        event.accept()

    def CopyImageToClipboard(self):
        current_image = self.image_panel.current_image
        if current_image is None:
            QMessageBox.information(self, "情報", "画像が読み込まれていません。")
            return
        if self.image_panel.file_name.lower().endswith((".jpg", ".jpeg")):
            QMessageBox.information(
                self,
                "コピーを中止しました",
                "容量が増えるためJPEG画像はクリップボードにコピーできません。\n"
                "挿入先から画像を取り込んでください。",
            )
            return
        try:
            QApplication.clipboard().setImage(pil_to_qimage(current_image))
            self.statusBar().showMessage("画像をクリップボードへコピーしました", 5000)
        except Exception as error:
            QMessageBox.critical(self, "コピーエラー", f"画像をコピーできませんでした。\n{error}")

    def PasteImageFromClipboard(self):
        try:
            clipboard = QApplication.clipboard()
            qimage = clipboard.image()
            if not qimage.isNull():
                self.image_panel.SetImage(qimage_to_pil(qimage), file_name="clipboard.png")
                self.image_panel.from_clipboard = True
                self.statusBar().showMessage("クリップボードから画像を貼り付けました", 5000)
                return

            # Qtで取得できない形式はPillowのWindows対応へフォールバックする
            pasted_image = ImageGrab.grabclipboard()
            if isinstance(pasted_image, Image.Image):
                self.image_panel.SetImage(pasted_image, file_name="clipboard.png")
                self.image_panel.from_clipboard = True
                self.statusBar().showMessage("クリップボードから画像を貼り付けました", 5000)
            else:
                QMessageBox.information(self, "情報", "クリップボードに画像がありません。")
        except Exception as error:
            QMessageBox.critical(
                self,
                "貼り付けエラー",
                f"クリップボードから画像を取得できませんでした。\n{error}",
            )


class ImageEditorApp(QApplication):
    """アプリ名と共通スタイルを設定するQApplication。"""

    def __init__(self, arguments):
        super().__init__(arguments)
        self.setApplicationName(APP_TITLE)
        self.setStyle("Fusion")
        self.setStyleSheet(APP_STYLE_SHEET)


def main():
    Image.MAX_IMAGE_PIXELS = 500_000_000
    app = ImageEditorApp(sys.argv)
    frame = ImageEditorFrame()
    frame.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
