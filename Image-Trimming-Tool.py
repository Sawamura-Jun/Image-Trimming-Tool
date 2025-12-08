import wx
import os
import io
import datetime
from PIL import Image, ImageDraw, ImageOps, ImageGrab  # ImageGrabでクリップボードからの取得を有効にする
# アプリケーションウィンドウの定数
APP_WINDOW_SIZE = (1224, 680)
DEFAULT_ROTATION_ANGLE = 0.1
DEFAULT_CROP_ASPECT = "1:1"
DEFAULT_IMAGE_SIZE = 1024
DEFAULT_JPEG_QUALITY = 70
LINES = 20
BACK_GROUND_COLOR = wx.Colour(100, 100, 100)
CLIPBOARD_SAVE_DIR = r""  # クリップボード保存先の上書き用。空のままならWindowsではPictures\\Image-Cropperを使用

def resolve_clipboard_save_dir():
    """
    クリップボード画像の保存先ディレクトリを返す。
    CLIPBOARD_SAVE_DIRが設定されていればそれを使い、未設定の場合はユーザープロファイル配下のPictures\\Image-Cropperを既定とする。
    """
    if CLIPBOARD_SAVE_DIR:
        return CLIPBOARD_SAVE_DIR
    home = os.path.expanduser("~")
    if home and home != "~":
        return os.path.join(home, "Pictures", "Image-Cropper")
    # ホームディレクトリが解決できないときはカレントディレクトリを使用
    return os.path.join(os.getcwd(), "Image-Cropper")

class ImagePanel(wx.Panel):
    HANDLE_SIZE = 10
    MIN_CROP_SIZE = 4

    def __init__(self, parent):
        super().__init__(parent)
        self.SetBackgroundColour(BACK_GROUND_COLOR)
        self.SetDoubleBuffered(True)
        self.Bind(wx.EVT_ERASE_BACKGROUND, self.OnEraseBackground)
        # スケーリングしたビットマップ描画用のキャッシュ
        self._cached_bitmap = None
        self._cached_size = (0, 0)
        self._cached_image_id = None
        self.original_image = None
        self.current_image = None
        self.file_name = ""
        # 現在の画像がクリップボードから取得された場合はTrue
        self.from_clipboard = False
        # トリミング矩形の状態と履歴を初期化
        self.crop_rect = None
        self.crop_history = []
        self.max_crop_history = 10
        self.mode = "idle"
        self.drag_handle = None
        self.drag_start = wx.Point()
        self.original_rect = None
        self.display_offset_x = 0
        self.display_offset_y = 0
        self.display_width = 0
        self.display_height = 0
        self.old_display_width = 0
        self.old_display_height = 0
        self.fixed_aspect = True
        self.Bind(wx.EVT_PAINT, self.OnPaint)
        self.Bind(wx.EVT_LEFT_DOWN, self.OnLeftDown)
        self.Bind(wx.EVT_LEFT_UP, self.OnLeftUp)
        self.Bind(wx.EVT_MOTION, self.OnMouseMove)
        self.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseLeave)
        self.Bind(wx.EVT_SIZE, self.OnResize)

    def OnEraseBackground(self, event):
        pass

    def OnResize(self, event):
        # 画面サイズ変更時に表示寸法を更新し、トリミング範囲を再スケーリング
        self.old_display_width = self.display_width
        self.old_display_height = self.display_height
        self.UpdateDisplayGeometry()
        self.RescaleCropRect()
        self.Refresh()
        event.Skip()

    def RescaleCropRect(self):
        # 直前の表示サイズが取得できない場合は処理しない
        if not self.crop_rect or self.old_display_width == 0 or self.old_display_height == 0:
            return
        old_x, old_y, old_w, old_h = self.crop_rect
        cx_old = old_x + old_w/2
        cy_old = old_y + old_h/2
        scale_x = 1.0
        scale_y = 1.0
        if self.old_display_width != 0:
            scale_x = self.display_width / self.old_display_width
        if self.old_display_height != 0:
            scale_y = self.display_height / self.old_display_height
        new_w = old_w * scale_x
        new_h = old_h * scale_y
        cx_new = cx_old * scale_x
        cy_new = cy_old * scale_y
        new_x = cx_new - new_w/2
        new_y = cy_new - new_h/2
        new_x, new_y, new_w, new_h = self.ClipRect(new_x, new_y, new_w, new_h)
        rect = self._ensure_min_size(wx.Rect(int(round(new_x)), int(round(new_y)), int(round(new_w)), int(round(new_h))))
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def ClipRect(self, x, y, w, h):
        # 矩形を表示エリア内に収める
        if w < 0:
            w = 0
        if h < 0:
            h = 0
        if x < 0:
            x = 0
        if x + w > self.display_width:
            x = self.display_width - w
        if y < 0:
            y = 0
        if y + h > self.display_height:
            y = self.display_height - h
        return (x, y, w, h)

    def UpdateDisplayGeometry(self):
        if not self.current_image:
            self.display_offset_x = 0
            self.display_offset_y = 0
            self.display_width = 0
            self.display_height = 0
            return
        panel_w, panel_h = self.GetClientSize()
        img_w, img_h = self.current_image.size
        scale = min(panel_w / img_w, panel_h / img_h)
        new_w = int(img_w * scale)
        new_h = int(img_h * scale)
        self.display_width, self.display_height = new_w, new_h
        self.display_offset_x = (panel_w - new_w) // 2
        self.display_offset_y = (panel_h - new_h) // 2
        # 表示サイズが変わったときはキャッシュをリセット
        self._cached_bitmap = None

    def _event_to_display_point(self, event):
        x, y = event.GetPosition()
        return wx.Point(x - self.display_offset_x, y - self.display_offset_y)

    def _clamp_display_point(self, point):
        return wx.Point(
            max(0, min(point.x, self.display_width)),
            max(0, min(point.y, self.display_height))
        )

    def _point_in_display(self, point):
        return 0 <= point.x <= self.display_width and 0 <= point.y <= self.display_height

    def _rect_contains_point(self, rect_tuple, point):
        x, y, w, h = rect_tuple
        return x <= point.x <= x + w and y <= point.y <= y + h

    def _iter_handle_rects_display(self):
        if not self.crop_rect:
            return []
        half = self.HANDLE_SIZE // 2
        x, y, w, h = self.crop_rect
        left = x
        top = y
        right = x + w
        bottom = y + h
        center_x = x + w / 2
        center_y = y + h / 2
        points = {
            "top_left": (left, top),
            "top": (center_x, top),
            "top_right": (right, top),
            "right": (right, center_y),
            "bottom_right": (right, bottom),
            "bottom": (center_x, bottom),
            "bottom_left": (left, bottom),
            "left": (left, center_y),
        }
        for name, (px, py) in points.items():
            yield name, wx.Rect(int(round(px)) - half, int(round(py)) - half, self.HANDLE_SIZE, self.HANDLE_SIZE)

    def _iter_handle_rects_panel(self):
        for name, rect in self._iter_handle_rects_display():
            rect.Offset(self.display_offset_x, self.display_offset_y)
            yield name, rect

    def _hit_test_handle(self, point):
        for name, rect in self._iter_handle_rects_display():
            if rect.Contains(point):
                return name
        return None

    def _get_aspect_ratio(self):
        if not self.fixed_aspect:
            return None
        try:
            aspect_str = getattr(self, "crop_aspect", DEFAULT_CROP_ASPECT)
            w_ratio, h_ratio = map(float, aspect_str.split(":"))
            if h_ratio == 0:
                return None
            return w_ratio / h_ratio
        except Exception:
            return None

    def _ensure_within_display(self, rect):
        if rect is None:
            return wx.Rect()
        rect = wx.Rect(rect)
        if self.display_width <= 0 or self.display_height <= 0:
            return wx.Rect()
        if rect.width > self.display_width:
            rect.width = self.display_width
        if rect.height > self.display_height:
            rect.height = self.display_height
        if rect.x < 0:
            rect.x = 0
        if rect.y < 0:
            rect.y = 0
        if rect.Right > self.display_width:
            rect.x = self.display_width - rect.width
        if rect.Bottom > self.display_height:
            rect.y = self.display_height - rect.height
        rect.width = max(self.MIN_CROP_SIZE, rect.width)
        rect.height = max(self.MIN_CROP_SIZE, rect.height)
        rect.x = max(0, min(rect.x, self.display_width - rect.width))
        rect.y = max(0, min(rect.y, self.display_height - rect.height))
        return rect

    def _ensure_min_size(self, rect):
        rect = self._ensure_within_display(rect)
        if rect.width < self.MIN_CROP_SIZE:
            rect.width = self.MIN_CROP_SIZE
        if rect.height < self.MIN_CROP_SIZE:
            rect.height = self.MIN_CROP_SIZE
        return self._ensure_within_display(rect)

    def _create_rect_with_ratio(self, anchor, current, ratio):
        dx = current.x - anchor.x
        dy = current.y - anchor.y
        abs_dx = abs(dx)
        abs_dy = abs(dy)
        if abs_dx == 0 and abs_dy == 0:
            return wx.Rect(anchor.x, anchor.y, 0, 0)
        if abs_dy == 0:
            abs_dy = int(round(abs_dx / ratio))
        if abs_dx == 0:
            abs_dx = int(round(abs_dy * ratio))
        current_ratio = abs_dx / abs_dy if abs_dy else ratio
        if current_ratio > ratio:
            abs_dx = int(round(abs_dy * ratio))
        else:
            abs_dy = int(round(abs_dx / ratio))
        x2 = anchor.x + (abs_dx if dx >= 0 else -abs_dx)
        y2 = anchor.y + (abs_dy if dy >= 0 else -abs_dy)
        left = min(anchor.x, x2)
        top = min(anchor.y, y2)
        rect = wx.Rect(left, top, abs(x2 - anchor.x), abs(y2 - anchor.y))
        return self._ensure_within_display(rect)

    def _create_rect(self, anchor, current):
        anchor = self._clamp_display_point(anchor)
        current = self._clamp_display_point(current)
        ratio = self._get_aspect_ratio()
        if ratio:
            rect = self._create_rect_with_ratio(anchor, current, ratio)
        else:
            left = min(anchor.x, current.x)
            top = min(anchor.y, current.y)
            width = abs(current.x - anchor.x)
            height = abs(current.y - anchor.y)
            rect = wx.Rect(left, top, width, height)
        return self._ensure_min_size(rect)

    def _rect_from_crop(self):
        if not self.crop_rect:
            return None
        x, y, w, h = self.crop_rect
        return wx.Rect(int(round(x)), int(round(y)), int(round(w)), int(round(h)))

    def _update_selection_creation(self, anchor, current):
        rect = self._create_rect(anchor, current)
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def _update_selection_move(self, dx, dy):
        if not self.original_rect:
            return
        rect = wx.Rect(self.original_rect)
        rect.x += dx
        rect.y += dy
        rect = self._ensure_min_size(rect)
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def _rect_from_horizontal_anchor(self, anchor_x, width, origin, to_left, ratio):
        width = max(self.MIN_CROP_SIZE, min(width, self.display_width))
        height = max(self.MIN_CROP_SIZE, int(round(width / ratio)))
        if height > self.display_height:
            height = self.display_height
            width = max(self.MIN_CROP_SIZE, int(round(height * ratio)))
        center_y = origin.y + origin.height / 2
        top = int(round(center_y - height / 2))
        top = max(0, min(top, self.display_height - height))
        bottom = top + height
        if to_left:
            right = min(anchor_x, self.display_width)
            left = max(0, right - width)
        else:
            left = max(0, anchor_x)
            right = min(self.display_width, left + width)
            left = right - width
        return wx.Rect(int(left), int(top), int(right - left), int(bottom - top))

    def _rect_from_vertical_anchor(self, anchor_y, height, origin, to_top, ratio):
        height = max(self.MIN_CROP_SIZE, min(height, self.display_height))
        width = max(self.MIN_CROP_SIZE, int(round(height * ratio)))
        if width > self.display_width:
            width = self.display_width
            height = max(self.MIN_CROP_SIZE, int(round(width / ratio)))
        center_x = origin.x + origin.width / 2
        left = int(round(center_x - width / 2))
        left = max(0, min(left, self.display_width - width))
        right = left + width
        if to_top:
            bottom = min(anchor_y, self.display_height)
            top = max(0, bottom - height)
        else:
            top = max(0, anchor_y)
            bottom = min(self.display_height, top + height)
            top = bottom - height
        return wx.Rect(int(left), int(top), int(right - left), int(bottom - top))

    def _resize_corner_with_ratio(self, point, handle, origin, ratio):
        if handle == "top_left":
            anchor = wx.Point(origin.Right, origin.Bottom)
            horizontal = -1
            vertical = -1
        elif handle == "top_right":
            anchor = wx.Point(origin.x, origin.Bottom)
            horizontal = 1
            vertical = -1
        elif handle == "bottom_left":
            anchor = wx.Point(origin.Right, origin.y)
            horizontal = -1
            vertical = 1
        else:
            anchor = wx.Point(origin.x, origin.y)
            horizontal = 1
            vertical = 1
        dx = (point.x - anchor.x) * horizontal
        dy = (point.y - anchor.y) * vertical
        dx = max(self.MIN_CROP_SIZE, min(abs(dx), self.display_width))
        dy = max(self.MIN_CROP_SIZE, min(abs(dy), self.display_height))
        if dy == 0:
            dy = int(round(dx / ratio))
        if dx == 0:
            dx = int(round(dy * ratio))
        width = dx
        height = int(round(width / ratio))
        if height > dy:
            height = dy
            width = int(round(height * ratio))
        height = max(self.MIN_CROP_SIZE, height)
        width = max(self.MIN_CROP_SIZE, width)
        if horizontal < 0:
            left = anchor.x - width
            right = anchor.x
        else:
            left = anchor.x
            right = anchor.x + width
        if vertical < 0:
            top = anchor.y - height
            bottom = anchor.y
        else:
            top = anchor.y
            bottom = anchor.y + height
        rect = wx.Rect(int(left), int(top), int(right - left), int(bottom - top))
        return self._ensure_min_size(rect)

    def _resize_with_ratio(self, point, handle, ratio):
        if not self.original_rect:
            return wx.Rect()
        origin = wx.Rect(self.original_rect)
        if handle == "left":
            anchor_x = origin.Right
            width = max(self.MIN_CROP_SIZE, min(anchor_x - point.x, anchor_x))
            rect = self._rect_from_horizontal_anchor(anchor_x, width, origin, to_left=True, ratio=ratio)
        elif handle == "right":
            anchor_x = origin.x
            width = max(self.MIN_CROP_SIZE, min(point.x - anchor_x, self.display_width - anchor_x))
            rect = self._rect_from_horizontal_anchor(anchor_x, width, origin, to_left=False, ratio=ratio)
        elif handle == "top":
            anchor_y = origin.Bottom
            height = max(self.MIN_CROP_SIZE, min(anchor_y - point.y, anchor_y))
            rect = self._rect_from_vertical_anchor(anchor_y, height, origin, to_top=True, ratio=ratio)
        elif handle == "bottom":
            anchor_y = origin.y
            height = max(self.MIN_CROP_SIZE, min(point.y - anchor_y, self.display_height - anchor_y))
            rect = self._rect_from_vertical_anchor(anchor_y, height, origin, to_top=False, ratio=ratio)
        else:
            rect = self._resize_corner_with_ratio(point, handle, origin, ratio)
        return self._ensure_min_size(rect)

    def _resize_free(self, point, handle):
        rect = wx.Rect(self.original_rect)
        left = rect.x
        top = rect.y
        right = rect.Right
        bottom = rect.Bottom
        if "left" in handle:
            left = min(point.x, right - self.MIN_CROP_SIZE)
        if "right" in handle:
            right = max(point.x, left + self.MIN_CROP_SIZE)
        if "top" in handle:
            top = min(point.y, bottom - self.MIN_CROP_SIZE)
        if "bottom" in handle:
            bottom = max(point.y, top + self.MIN_CROP_SIZE)
        left = max(0, min(left, self.display_width))
        right = max(0, min(right, self.display_width))
        top = max(0, min(top, self.display_height))
        bottom = max(0, min(bottom, self.display_height))
        width = max(self.MIN_CROP_SIZE, right - left)
        height = max(self.MIN_CROP_SIZE, bottom - top)
        return wx.Rect(int(left), int(top), int(width), int(height))

    def _update_selection_resize(self, point):
        if not self.original_rect or not self.drag_handle:
            return
        ratio = self._get_aspect_ratio() if self.fixed_aspect else None
        if ratio:
            rect = self._resize_with_ratio(point, self.drag_handle, ratio)
        else:
            rect = self._resize_free(point, self.drag_handle)
        rect = self._ensure_min_size(rect)
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def _update_cursor(self, display_point):
        handle = self._hit_test_handle(display_point)
        if handle in ("top_left", "bottom_right"):
            cursor = wx.CURSOR_SIZENWSE
        elif handle in ("top_right", "bottom_left"):
            cursor = wx.CURSOR_SIZENESW
        elif handle in ("left", "right"):
            cursor = wx.CURSOR_SIZEWE
        elif handle in ("top", "bottom"):
            cursor = wx.CURSOR_SIZENS
        elif self._point_in_display(display_point):
            if self.crop_rect and self._rect_contains_point(self.crop_rect, display_point):
                cursor = wx.CURSOR_SIZING
            else:
                cursor = wx.CURSOR_CROSS
        else:
            cursor = wx.CURSOR_ARROW
        self.SetCursor(wx.Cursor(cursor))

    def ApplyAspectRatioToSelection(self):
        if not self.fixed_aspect or not self.crop_rect:
            return
        ratio = self._get_aspect_ratio()
        if not ratio:
            return
        x, y, w, h = self.crop_rect
        center_x = x + w / 2
        center_y = y + h / 2
        width = w
        height = h
        desired_height = int(round(width / ratio))
        desired_width = int(round(height * ratio))
        if desired_height <= self.display_height:
            height = desired_height
        if height <= 0:
            height = self.MIN_CROP_SIZE
        width = int(round(height * ratio))
        if width > self.display_width:
            width = self.display_width
            height = int(round(width / ratio))
        if height > self.display_height:
            height = self.display_height
            width = int(round(height * ratio))
        width = max(self.MIN_CROP_SIZE, width)
        height = max(self.MIN_CROP_SIZE, height)
        left = int(round(center_x - width / 2))
        top = int(round(center_y - height / 2))
        rect = self._ensure_min_size(wx.Rect(left, top, width, height))
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)

    def SetImage(self, pil_image, file_name=""):
        # ディスクから読み込むときはクリップボードフラグをリセット
        self.from_clipboard = False
        self.original_image = pil_image.copy()
        self.current_image = pil_image.copy()
        self.crop_history = [self.current_image.copy()]
        self.mode = "idle"
        self.drag_handle = None
        self.original_rect = None
        self.drag_start = wx.Point()
        # 新しい画像を読み込んだあとに回転の基準をリセット
        self.rotation_base_image = self.current_image.copy()
        self.rotation_angle_total = 0.0
        self.file_name = os.path.basename(file_name)
        self.file_dir = os.path.dirname(file_name)
        self.UpdateDisplayGeometry()
        self.InitCropRect()
        self.UpdateTitle()
        self._cached_bitmap = None
        self.Refresh()

    def UpdateTitle(self):
        if self.file_name and self.current_image:
            w, h = self.current_image.size
            title = f"{self.file_name} ({w}x{h})"
            top_frame = self.GetTopLevelParent()
            if top_frame:
                top_frame.SetTitle(title)

    def OnPaint(self, event):
        dc = wx.BufferedPaintDC(self)
        dc.Clear()
        if self.current_image:
            pos_x = self.display_offset_x
            pos_y = self.display_offset_y
            # 画像やサイズが変わったときにキャッシュを作り直す
            if (self._cached_bitmap is None or
                self._cached_size != (self.display_width, self.display_height) or
                self._cached_image_id != id(self.current_image)):
                # インタラクティブな再描画にはバイリニア補間を使用
                img_tmp = self.current_image.resize((self.display_width, self.display_height), Image.BILINEAR)
                buf = img_tmp.convert("RGB").tobytes()
                self._cached_bitmap = wx.Bitmap.FromBuffer(self.display_width, self.display_height, buf)
                self._cached_size = (self.display_width, self.display_height)
                self._cached_image_id = id(self.current_image)
            dc.DrawBitmap(self._cached_bitmap, pos_x, pos_y)
            # ガイドラインのグリッドを描画
            gc = wx.GraphicsContext.Create(dc)
            if gc:
                pen = wx.Pen(wx.Colour(255,255,255), width=1, style=wx.PENSTYLE_DOT)
                gc.SetPen(pen)
                for i in range(LINES+1):
                    yy = pos_y + int(self.display_height * i / LINES)
                    gc.StrokeLine(pos_x, yy, pos_x + self.display_width, yy)
                    xx = pos_x + int(self.display_width * i / LINES)
                    gc.StrokeLine(xx, pos_y, xx, pos_y + self.display_height)
            # トリミング範囲のオーバーレイを描画
            if self.crop_rect:
                crop_x = self.crop_rect[0] + pos_x
                crop_y = self.crop_rect[1] + pos_y
                crop_w = self.crop_rect[2]
                crop_h = self.crop_rect[3]
                if gc:
                    overlay_path = gc.CreatePath()
                    panel_w, panel_h = self.GetClientSize()
                    overlay_path.AddRectangle(0, 0, panel_w, panel_h)
                    rect_path = gc.CreatePath()
                    rect_path.AddRectangle(crop_x, crop_y, crop_w, crop_h)
                    overlay_path.AddPath(rect_path)
                    overlay_path.CloseSubpath()
                    gc.SetBrush(wx.Brush(wx.Colour(0, 0, 0, 100), wx.BRUSHSTYLE_SOLID))     # 100 = overlay alpha
                    gc.FillPath(overlay_path, wx.ODDEVEN_RULE)
                    gc.SetPen(wx.Pen(wx.Colour(255, 0, 0), 1, wx.PENSTYLE_SOLID))
                    gc.StrokePath(rect_path)
                try:
                    target_dc = wx.GCDC(dc)
                except Exception:
                    target_dc = dc
                target_dc.SetPen(wx.Pen(wx.Colour(255, 0, 0), 1))
                target_dc.SetBrush(wx.Brush(wx.Colour(255, 255, 255)))
                for _, handle_rect in self._iter_handle_rects_panel():
                    target_dc.DrawRectangle(handle_rect)

    def OnLeftUp(self, event):
        previous_mode = self.mode
        if self.HasCapture():
            self.ReleaseMouse()
        self.mode = "idle"
        self.drag_handle = None
        self.original_rect = None
        self.drag_start = wx.Point()
        if previous_mode == "creating" and self.crop_rect:
            rect = self._ensure_min_size(self._rect_from_crop())
            self.crop_rect = (rect.x, rect.y, rect.width, rect.height)
        self._update_cursor(self._event_to_display_point(event))
        event.Skip()

    def OnMouseMove(self, event):
        if not self.current_image:
            return
        display_point = self._event_to_display_point(event)
        if event.Dragging() and event.LeftIsDown() and self.mode != "idle":
            point = self._clamp_display_point(display_point)
            if self.mode == "creating":
                self._update_selection_creation(self.drag_start, point)
            elif self.mode == "moving" and self.original_rect:
                dx = point.x - self.drag_start.x
                dy = point.y - self.drag_start.y
                self._update_selection_move(dx, dy)
            elif self.mode == "resizing" and self.original_rect:
                self._update_selection_resize(point)
            self.Refresh(False)
            return
        self._update_cursor(display_point)

    def OnMouseLeave(self, event):
        if self.mode == "idle":
            self.SetCursor(wx.Cursor(wx.CURSOR_ARROW))

    def OnLeftDown(self, event):
        if not self.current_image:
            return
        display_point = self._event_to_display_point(event)
        handle = self._hit_test_handle(display_point)
        if handle and self.crop_rect:
            self.mode = "resizing"
            self.drag_handle = handle
            self.original_rect = self._rect_from_crop()
            self.drag_start = self._clamp_display_point(display_point)
        elif self.crop_rect and self._rect_contains_point(self.crop_rect, display_point):
            self.mode = "moving"
            self.drag_handle = "inside"
            self.original_rect = self._rect_from_crop()
            self.drag_start = self._clamp_display_point(display_point)
        elif self._point_in_display(display_point):
            self.mode = "creating"
            self.drag_handle = None
            anchor = self._clamp_display_point(display_point)
            self.drag_start = anchor
            self.original_rect = None
            self.crop_rect = (anchor.x, anchor.y, 0, 0)
        else:
            self.mode = "idle"
            self.drag_handle = None
            self.original_rect = None
            return
        if not self.HasCapture():
            self.CaptureMouse()
        self.Refresh(False)

    def RotateImage(self, delta):
        if self.rotation_base_image:
            # 累積回転角を0〜360度の範囲に保つ
            self.rotation_angle_total = (self.rotation_angle_total + delta) % 360
            rotated = self.rotation_base_image.rotate(self.rotation_angle_total, expand=True, resample=Image.BICUBIC)
            self.current_image = rotated
            if len(self.crop_history) >= self.max_crop_history:
                self.crop_history.pop(0)
            self.crop_history.append(self.current_image.copy())
            self.UpdateDisplayGeometry()
            self.UpdateTitle()
            self.Refresh()

    def CropImage(self):
        if self.crop_rect and self.current_image:
            pos_x = self.display_offset_x
            pos_y = self.display_offset_y
            disp_w = self.display_width
            disp_h = self.display_height
            img_w, img_h = self.current_image.size
            crop_x = max(0, self.crop_rect[0])
            crop_y = max(0, self.crop_rect[1])
            if disp_w == 0 or disp_h == 0:
                return
            scale_x = img_w / disp_w
            scale_y = img_h / disp_h
            rect_w = self.crop_rect[2]
            rect_h = self.crop_rect[3]
            # 表示座標をround()を使って画像座標へ変換
            x = round(crop_x * scale_x)
            y = round(crop_y * scale_y)
            w = round(rect_w * scale_x)
            h = round(rect_h * scale_y)
            if w == 0 or h == 0:
                return
            cropped = self.current_image.crop((x, y, x + w, y + h))
            if len(self.crop_history) >= self.max_crop_history:
                self.crop_history.pop(0)
            self.current_image = cropped
            self.crop_history.append(self.current_image.copy())
            self.UpdateDisplayGeometry()
            self.InitCropRect()
            self.UpdateTitle()
            self.Refresh()
            # トリミング後に回転の基準をリセット
            self.rotation_base_image = self.current_image.copy()
            self.rotation_angle_total = 0.0

    def RevertCrop(self):
        if len(self.crop_history) > 1:
            self.crop_history.pop()
            self.current_image = self.crop_history[-1].copy()
            self.UpdateDisplayGeometry()
            # 現在のファイル名とサイズでタイトルを更新
            self.InitCropRect()
            self.UpdateTitle()
            self.Refresh()
            # 更新された表示サイズを反映させるために再描画
            self.rotation_base_image = self.current_image.copy()
            self.rotation_angle_total = 0.0

    def ResizeImage(self, target_size):
        if self.current_image:
            w, h = self.current_image.size
            long_side = max(w, h)
            if long_side > target_size:
                ratio = target_size / long_side
                new_w = int(w * ratio)
                new_h = int(h * ratio)
                resized = self.current_image.resize((new_w, new_h), Image.LANCZOS)
                if len(self.crop_history) >= self.max_crop_history:
                    self.crop_history.pop(0)
                self.current_image = resized
                self.crop_history.append(self.current_image.copy())
                self.UpdateDisplayGeometry()
                self.InitCropRect()
                self.UpdateTitle()
                self.Refresh()
                # 画像サイズ変更後に回転の基準をリセット
                self.rotation_base_image = self.current_image.copy()
                self.rotation_angle_total = 0.0

    def SaveImage(self, jpeg_quality):
        if self.current_image:
            if self.from_clipboard:
                # クリップボードからの画像はPNGでタイムスタンプ付き保存
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                file_name = f"clipboard_{timestamp}.png"
                save_dir = resolve_clipboard_save_dir()
                os.makedirs(save_dir, exist_ok=True)
                save_path = os.path.join(save_dir, file_name)
                self.current_image.save(save_path, "PNG")
            elif self.file_name:
                name, ext = os.path.splitext(self.file_name)
                new_name = name + "_trm" + ext
                save_path = os.path.join(self.file_dir, new_name)
                params = {}
                if ext.lower() in [".jpg", ".jpeg"]:
                    params["quality"] = jpeg_quality
                self.current_image.save(save_path, **params)

    def InitCropRect(self):
        disp_w = self.display_width
        disp_h = self.display_height
        try:
            # 固定時はテキストボックスの縦横比を使用
            if self.fixed_aspect:
                aspect_str = getattr(self, "crop_aspect", DEFAULT_CROP_ASPECT)
                w_ratio, h_ratio = map(float, aspect_str.split(":"))
                if w_ratio < h_ratio:
                    rect_h = disp_h // 4
                    rect_w = int(rect_h * (w_ratio / h_ratio))
                else:
                    rect_w = disp_w // 4
                    rect_h = int(rect_w * (h_ratio / w_ratio)) if w_ratio != 0 else disp_h // 4
            else:
                # 解析に失敗した場合は1/4サイズの正方形をデフォルトにする
                rect_w = disp_w // 4
                rect_h = rect_w
        except Exception:
            rect_w = disp_w // 4
            rect_h = rect_w
        center_x = disp_w / 2
        center_y = disp_h / 2
        new_x = center_x - rect_w / 2
        new_y = center_y - rect_h / 2
        rect = wx.Rect(int(round(new_x)), int(round(new_y)), int(round(rect_w)), int(round(rect_h)))
        rect = self._ensure_min_size(rect)
        self.crop_rect = (rect.x, rect.y, rect.width, rect.height)
        self.ApplyAspectRatioToSelection()
        self.mode = "idle"
        self.drag_handle = None
        self.original_rect = None
        self.drag_start = wx.Point()

class ControlPanel(wx.Panel):
    def __init__(self, parent, image_panel):
        super().__init__(parent, size=(250, -1))
        self.image_panel = image_panel
        self.InitUI()

    def InitUI(self):
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add((0, 20), 0, wx.EXPAND)
        font = wx.Font(17, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        # 回転
        hbox_rot = wx.BoxSizer(wx.HORIZONTAL)
        st_rot = wx.StaticText(self, label="回転角度:")
        st_rot.SetFont(font)
        self.tc_rot = wx.TextCtrl(self, value=str(DEFAULT_ROTATION_ANGLE), style=wx.TE_CENTER, size=(100,35))
        self.tc_rot.SetFont(font)
        hbox_rot.Add(st_rot, flag=wx.RIGHT, border=5)
        hbox_rot.Add(self.tc_rot, proportion=1)
        vbox.Add(hbox_rot, flag=wx.EXPAND | wx.ALL, border=5)
        self.current_angle = 0.0
        hbox_rot_btn = wx.BoxSizer(wx.HORIZONTAL)
        btn_rot_left = wx.Button(self, label="左回転", size=(100,45))
        btn_rot_left.SetFont(font)
        btn_rot_right = wx.Button(self, label="右回転", size=(100,45))
        btn_rot_right.SetFont(font)
        btn_rot_left.Bind(wx.EVT_BUTTON, self.OnRotateLeft)
        btn_rot_right.Bind(wx.EVT_BUTTON, self.OnRotateRight)
        hbox_rot_btn.Add(btn_rot_left, proportion=1, flag=wx.RIGHT, border=5)
        hbox_rot_btn.Add(btn_rot_right, proportion=1)
        vbox.Add(hbox_rot_btn, flag=wx.EXPAND | wx.ALL, border=5)
        vbox.Add((0, 50), 0, wx.EXPAND)
        # 縦横比
        hbox_crop = wx.BoxSizer(wx.HORIZONTAL)
        self.cb_aspect = wx.CheckBox(self, label="")
        self.cb_aspect.SetValue(True)
        self.cb_aspect.Bind(wx.EVT_CHECKBOX, self.OnAspectCheckbox)
        hbox_crop.Add(self.cb_aspect, flag=wx.RIGHT, border=5)
        st_crop = wx.StaticText(self, label="縦横比")
        st_crop.SetFont(font)
        hbox_crop.Add(st_crop, flag=wx.RIGHT, border=5)
        self.tc_crop = wx.TextCtrl(self, value=DEFAULT_CROP_ASPECT, style=wx.TE_PROCESS_ENTER | wx.TE_CENTER, size=(100,35))
        self.tc_crop.SetFont(font)
        self.tc_crop.Bind(wx.EVT_TEXT_ENTER, self.OnAspectEnter)
        hbox_crop.Add(self.tc_crop, proportion=1)
        vbox.Add(hbox_crop, flag=wx.EXPAND | wx.ALL, border=5)
        btn_crop = wx.Button(self, label="トリミング", size=(100,45))
        btn_crop.SetFont(font)
        btn_crop.Bind(wx.EVT_BUTTON, self.OnCrop)
        vbox.Add(btn_crop, flag=wx.EXPAND | wx.ALL, border=5)
        btn_revert = wx.Button(self, label="もどる", size=(100,45))
        btn_revert.SetFont(font)
        btn_revert.Bind(wx.EVT_BUTTON, self.OnRevert)
        vbox.Add(btn_revert, flag=wx.EXPAND | wx.ALL, border=5)
        vbox.Add((0, 50), 0, wx.EXPAND)
        # サイズ
        hbox_resize = wx.BoxSizer(wx.HORIZONTAL)
        st_resize = wx.StaticText(self, label="画像サイズ:")
        st_resize.SetFont(font)
        self.tc_resize = wx.TextCtrl(self, value=str(DEFAULT_IMAGE_SIZE), style=wx.TE_CENTER, size=(100,35))
        self.tc_resize.SetFont(font)
        hbox_resize.Add(st_resize, flag=wx.RIGHT, border=5)
        hbox_resize.Add(self.tc_resize, proportion=1)
        vbox.Add(hbox_resize, flag=wx.EXPAND | wx.ALL, border=5)
        btn_resize = wx.Button(self, label="画像サイズ変更", size=(100,45))
        btn_resize.SetFont(font)
        btn_resize.Bind(wx.EVT_BUTTON, self.OnResizeImage)
        vbox.Add(btn_resize, flag=wx.EXPAND | wx.ALL, border=5)
        vbox.Add((0, 50), 0, wx.EXPAND)
        # 保存
        hbox_quality = wx.BoxSizer(wx.HORIZONTAL)
        st_quality = wx.StaticText(self, label="JPG品質")
        st_quality.SetFont(font)
        self.tc_quality = wx.TextCtrl(self, value=str(DEFAULT_JPEG_QUALITY), style=wx.TE_CENTER, size=(100,35))
        self.tc_quality.SetFont(font)
        hbox_quality.Add(st_quality, flag=wx.RIGHT, border=5)
        hbox_quality.Add(self.tc_quality, proportion=1)
        vbox.Add(hbox_quality, flag=wx.EXPAND | wx.ALL, border=5)
        btn_save = wx.Button(self, label="保存", size=(100,45))
        btn_save.SetFont(font)
        btn_save.Bind(wx.EVT_BUTTON, self.OnSave)
        vbox.Add(btn_save, flag=wx.EXPAND | wx.ALL, border=5)
        self.SetSizer(vbox)

    def OnRotateLeft(self, event):
        try:
            delta = float(self.tc_rot.GetValue())
            self.image_panel.RotateImage(delta)
        except ValueError:
            wx.MessageBox("回転角度に数値を入力してください。", "エラー", wx.OK | wx.ICON_ERROR)

    def OnRotateRight(self, event):
        try:
            delta = float(self.tc_rot.GetValue())
            self.image_panel.RotateImage(-delta)
        except ValueError:
            wx.MessageBox("回転角度に数値を入力してください。", "エラー", wx.OK | wx.ICON_ERROR)

    def OnAspectEnter(self, event):
        ratio_str = self.tc_crop.GetValue()
        try:
            w_ratio, h_ratio = map(float, ratio_str.split(":"))
            if h_ratio == 0:
                raise ValueError
            self.image_panel.crop_aspect = ratio_str
            self.image_panel.fixed_aspect = self.cb_aspect.GetValue()
            if self.image_panel.crop_rect:
                self.image_panel.ApplyAspectRatioToSelection()
            else:
                self.image_panel.InitCropRect()
            self.image_panel.Refresh()
        except Exception:
            wx.MessageBox("縦横比の入力形式が不正です。例: 1:1", "エラー", wx.OK | wx.ICON_ERROR)

    def OnAspectCheckbox(self, event):
        self.image_panel.fixed_aspect = self.cb_aspect.GetValue()
        # チェックボックスがオンのときは指定の縦横比を維持
        if self.cb_aspect.GetValue():
            if self.image_panel.crop_rect:
                self.image_panel.ApplyAspectRatioToSelection()
            else:
                self.image_panel.InitCropRect()
        self.image_panel.Refresh()

    def OnCrop(self, event):
        """トリミングボタン押下時の処理。"""
        if not self.image_panel.crop_rect:
            return
        if self.cb_aspect.GetValue():
            ratio_str = self.tc_crop.GetValue()
            try:
                w_ratio, h_ratio = map(float, ratio_str.split(":"))
                old_x, old_y, old_w, old_h = self.image_panel.crop_rect
                new_w = old_w
                new_h = round(new_w * (h_ratio / w_ratio)) if w_ratio != 0 else old_h
                cx = old_x + old_w / 2
                cy = old_y + old_h / 2
                new_x = cx - new_w / 2
                new_y = cy - new_h / 2
                self.image_panel.crop_rect = (new_x, new_y, new_w, new_h)
                self.image_panel.Refresh()
            except Exception:
                wx.MessageBox("縦横比の入力形式が不正です。例: 1:1", "エラー", wx.OK | wx.ICON_ERROR)
                return
        # 現在のトリミング範囲を画像に反映
        self.image_panel.CropImage()

    def OnRevert(self, event):
        self.image_panel.RevertCrop()

    def OnResizeImage(self, event):
        try:
            target_size = int(self.tc_resize.GetValue())
            self.image_panel.ResizeImage(target_size)
        except ValueError:
            wx.MessageBox("画像サイズに数値を入力してください。", "エラー", wx.OK | wx.ICON_ERROR)

    def OnSave(self, event):
        try:
            quality = int(self.tc_quality.GetValue())
            self.image_panel.SaveImage(quality)
        except ValueError:
            wx.MessageBox("圧縮率に数値を入力してください。", "エラー", wx.OK | wx.ICON_ERROR)

class FileDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        super().__init__()
        self.window = window

    def OnDropFiles(self, x, y, filenames):
        if filenames:
            try:
                img = Image.open(filenames[0])
                self.window.image_panel.SetImage(img, file_name=filenames[0])
            except Exception:
                wx.MessageBox("画像ファイルの読み込みに失敗しました。", "エラー", wx.OK | wx.ICON_ERROR)
        return True

class ImageEditorFrame(wx.Frame):

    def __init__(self):
        super().__init__(None, title="Image-Cropper", size=APP_WINDOW_SIZE)
        self.InitUI()
        self.Centre()
        self.Show()
        self.Refresh()  # 描画欠けを防ぐため初回に再描画を強制
        self.Update()
        # グローバルショートカットキー（例: Ctrl+V）をバインド
        self.Bind(wx.EVT_CHAR_HOOK, self.OnKeyDown)

    def InitUI(self):
        panel = wx.Panel(self)
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.image_panel = ImagePanel(panel)
        control_panel = ControlPanel(panel, self.image_panel)
        hbox.Add(self.image_panel, proportion=1, flag=wx.EXPAND)
        hbox.Add(control_panel, proportion=0, flag=wx.EXPAND)
        panel.SetSizer(hbox)
        dt = FileDropTarget(self)
        self.SetDropTarget(dt)

    def OnKeyDown(self, event):
        keycode = event.GetKeyCode()
        if event.ControlDown() and keycode == ord('V'):
            self.PasteImageFromClipboard()
        elif event.ControlDown() and keycode == ord('C'):
            self.CopyImageToClipboard()
        else:
            event.Skip()

    def CopyImageToClipboard(self):
        current_image = self.image_panel.current_image
        if not current_image:
            wx.MessageBox("画像が読み込まれていません。", "情報", wx.OK | wx.ICON_INFORMATION)
            return
        if self.image_panel.file_name.lower().endswith((".jpg", ".jpeg")):
            wx.MessageBox("容量が増えるためJpeg画像はクリップボードにコピーできません\n挿入から画像を取り込んでください", "コピーを中止しました", wx.OK | wx.ICON_INFORMATION)
            return
        try:
            buffer = io.BytesIO()
            current_image.save(buffer, format="PNG")
            png_bytes = buffer.getvalue()
            data = wx.CustomDataObject(wx.DataFormat("PNG"))
            data.SetData(png_bytes)
            if wx.TheClipboard.Open():
                wx.TheClipboard.SetData(data)
                wx.TheClipboard.Flush()
                wx.TheClipboard.Close()
            else:
                wx.MessageBox("クリップボードを開けませんでした。", "エラー", wx.OK | wx.ICON_ERROR)
        except Exception:
            wx.MessageBox("画像のコピーに失敗しました。", "エラー", wx.OK | wx.ICON_ERROR)

    def PasteImageFromClipboard(self):
        try:
            # PIL.ImageGrab.grabclipboard()でクリップボードの画像データを取得
            pasted_image = ImageGrab.grabclipboard()
            if pasted_image:
                # クリップボード画像をタイムスタンプ付きPNG名で保存し、フラグをオンにする
                self.image_panel.SetImage(pasted_image, file_name="clipboard.png")
                self.image_panel.from_clipboard = True
            else:
                wx.MessageBox("クリップボードに画像がありません。", "情報", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox("クリップボードからの画像取得に失敗しました。", "エラー", wx.OK | wx.ICON_ERROR)

class ImageEditorApp(wx.App):
    def OnInit(self):
        self.frame = ImageEditorFrame()
        self.SetTopWindow(self.frame)
        return True

if __name__ == "__main__":
    Image.MAX_IMAGE_PIXELS = 500000000
    app = ImageEditorApp(False)
    app.MainLoop()
