import importlib.util
import os
from pathlib import Path

import pytest
from PIL import Image
from PySide6.QtWidgets import QApplication


# 画面を表示しない環境でもQtウィジェットを検証できるようにする
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

MODULE_PATH = Path(__file__).with_name("Image-Trimming-Tool.py")
SPEC = importlib.util.spec_from_file_location("image_trimming_tool", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


@pytest.fixture(scope="module")
def app():
    application = QApplication.instance() or MODULE.ImageEditorApp([])
    yield application


def test_parse_aspect_ratio():
    assert MODULE.parse_aspect_ratio("16:9") == pytest.approx(16 / 9)
    with pytest.raises(ValueError):
        MODULE.parse_aspect_ratio("1:0")
    with pytest.raises(ValueError):
        MODULE.parse_aspect_ratio("invalid")


def test_qimage_round_trip_keeps_size_and_color():
    source = Image.new("RGBA", (7, 5), (12, 34, 56, 200))
    restored = MODULE.qimage_to_pil(MODULE.pil_to_qimage(source))
    assert restored.size == source.size
    assert restored.getpixel((0, 0)) == source.getpixel((0, 0))


def test_image_panel_crop_resize_rotate_and_revert(app):
    panel = MODULE.ImagePanel()
    panel.resize(800, 500)
    panel.show()
    app.processEvents()

    source = Image.new("RGB", (400, 200), "navy")
    panel.SetImage(source, "sample.png")
    assert panel.display_width > 0
    assert panel.crop_rect is not None

    # 表示領域の中央半分を切り出す
    panel.fixed_aspect = False
    panel.crop_rect = (
        panel.display_width // 4,
        panel.display_height // 4,
        panel.display_width // 2,
        panel.display_height // 2,
    )
    panel.CropImage()
    assert panel.current_image.size == (200, 100)

    panel.ResizeImage(100)
    assert panel.current_image.size == (100, 50)

    panel.RotateImage(90)
    assert panel.current_image.size == (50, 100)

    panel.RevertCrop()
    assert panel.current_image.size == (100, 50)
    panel.close()


def test_main_window_builds_without_layout_errors(app):
    window = MODULE.ImageEditorFrame()
    window.show()
    app.processEvents()
    assert window.image_panel.width() >= window.image_panel.minimumWidth()
    assert window.control_panel.width() == MODULE.CONTROL_PANEL_WIDTH
    assert window.minimumSize().toTuple() == MODULE.APP_WINDOW_SIZE
    assert window.statusBar().currentMessage()
    window.close()
