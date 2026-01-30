import sys
import os
import math
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
from dataclasses import dataclass
import json
import numpy as np
from PIL import Image
import cv2
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import io

@dataclass
class AppConfig:
	"""应用程序配置类"""
	window_width: int = 1000
	window_height: int = 700
	selection_x_normalized: float = 0.1  # 框相对图片的xywh
	selection_y_normalized: float = 0.1
	selection_w_normalized: float = 0.8
	selection_h_normalized: float = 0.8
	image_scale: float = 1.  # 默认填满当前区域大小前提下进行的缩放倍数
	# image_translation_x_normalized: float = 0.  # 相对平移
	# image_translation_y_normalized: float = 0.  # 相对平移
	grid_rows: int = 3
	grid_cols: int = 3
	cut_border: bool = False
	preview_mode: bool = False
	pdf_preset: str = 'A4'
	pdf_width_spin: float = 21.0  # 默认A4宽度
	pdf_height_spin: float = 29.7  # 默认A4高度
	@classmethod
	def load(cls, filename="config.json"):
		"""从文件加载配置"""
		try:
			if os.path.exists(filename):
				with open(filename, 'r') as f:
					data = json.load(f)
				return cls(**data)
		except:
			pass
		return cls()

	def save(self, filename="config.json"):
		"""保存配置到文件"""
		try:
			with open(filename, 'w') as f:
				json.dump({
					'window_width': self.window_width,
					'window_height': self.window_height,
					'selection_x_normalized': self.selection_x_normalized,
					'selection_y_normalized': self.selection_y_normalized,
					'selection_w_normalized': self.selection_w_normalized,
					'selection_h_normalized': self.selection_h_normalized,
					'image_scale': self.image_scale,
					# 'image_translation_x_normalized': self.image_translation_x_normalized,
					# 'image_translation_y_normalized': self.image_translation_y_normalized,
					'grid_rows': self.grid_rows,
					'grid_cols': self.grid_cols,
					'cut_border': self.cut_border,
					'preview_mode': self.preview_mode,
					'pdf_preset': self.pdf_preset,
					'pdf_width_spin': self.pdf_width_spin,
					'pdf_height_spin': self.pdf_height_spin,
				}, f, indent=2)
		except:
			pass

class DraggableSelectionBox(QGraphicsView):
	"""可拖拽调整的选区框 (QGraphicsView)"""

	# 定义调整区域大小
	HANDLE_SIZE = 8

	# 定义调整区域类型
	NO_ADJUST = 0
	TOP_LEFT = 1
	TOP_RIGHT = 2
	BOTTOM_LEFT = 3
	BOTTOM_RIGHT = 4
	TOP = 5
	BOTTOM = 6
	LEFT = 7
	RIGHT = 8
	MOVE = 9

	def __init__(self, parent: ImageGridSplitter=None):
		super().__init__()
		self._parent = parent
		self.selection_rect = QRectF(10, 10, 10, 10)  # 默认选区
		self.is_dragging = False
		self.drag_start_pos = QPointF()
		self.drag_type = self.NO_ADJUST
		self.original_rect = QRectF()
		self.is_panning = False
		self.pan_start_pos = QPoint()
		self._has_initial_fit = False
		
		self.setMouseTracking(True)
		self.setRenderHint(QPainter.Antialiasing)
		self.setFrameStyle(QFrame.Shape.NoFrame)
		self.setBackgroundBrush(QBrush(QColor(43, 43, 43)))
		self.setTransformationAnchor(QGraphicsView.ViewportAnchor.NoAnchor)
		self.setResizeAnchor(QGraphicsView.ViewportAnchor.NoAnchor)
		
		self.scene = QGraphicsScene(self)
		self.setScene(self.scene)
		
		self.image_item = QGraphicsPixmapItem()
		self.image_item.hide()  # 初始隐藏
		self.image_item.setZValue(0)
		self.scene.addItem(self.image_item)
		
		self.selection_item = QGraphicsRectItem()
		self.selection_item.hide()  # 初始隐藏
		self.selection_item.setZValue(10)
		self.selection_item.setPen(QPen(QColor(0, 120, 215), 2))
		self.selection_item.setBrush(Qt.BrushStyle.NoBrush)
		self.scene.addItem(self.selection_item)
		
		self.handle_items = []
		for handle_type in (self.TOP_LEFT, self.TOP_RIGHT, self.BOTTOM_LEFT, self.BOTTOM_RIGHT):
			handle = QGraphicsEllipseItem()
			handle.hide()  # 初始隐藏
			handle.setRect(-self.HANDLE_SIZE/2, -self.HANDLE_SIZE/2, self.HANDLE_SIZE, self.HANDLE_SIZE)
			handle.setBrush(QBrush(QColor(0, 120, 215)))
			handle.setPen(QPen(QColor(0, 120, 215)))
			handle.setZValue(11)
			handle.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIgnoresTransformations, True)
			handle.setData(0, handle_type)
			self.scene.addItem(handle)
			self.handle_items.append(handle)
		
		self.grid_items = []
		self.preview_items = []

	def set_pixmap(self, pixmap: QPixmap, apply_fit: bool = True):
		self.image_item.setPixmap(pixmap)
		self.scene.setSceneRect(QRectF(0, 0, pixmap.width(), pixmap.height()))
		if apply_fit:
			self.resetTransform()
			self.fitInView(self.scene.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)
			if self._parent:
				scale = self._parent.config.image_scale
				if scale != 1.0:
					self.scale(scale, scale)
		self._has_initial_fit = True

		self.image_item.show()
		self.selection_item.show()
		for handle in self.handle_items:
			handle.show()

	def set_selection_rect(self, rect: QRectF, rows: int, cols: int):
		self.selection_rect = rect
		self._update_selection_items()
		self.update_grid_items(rows, cols)

	def _update_selection_items(self):
		self.selection_item.setRect(self.selection_rect)
		corners = [
			self.selection_rect.topLeft(),
			self.selection_rect.topRight(),
			self.selection_rect.bottomLeft(),
			self.selection_rect.bottomRight()
		]
		for handle, pos in zip(self.handle_items, corners):
			handle.setPos(pos)

	def update_grid_items(self, rows: int, cols: int):
		for item in self.grid_items:
			self.scene.removeItem(item)
		self.grid_items = []
		
		if rows <= 1 and cols <= 1:
			return
		
		rect = self.selection_rect
		cell_width = rect.width() / cols
		cell_height = rect.height() / rows
		pen = QPen(QColor(0, 120, 215), 2, Qt.PenStyle.DashLine)
		
		for i in range(1, cols):
			x = rect.left() + i * cell_width
			line = QGraphicsLineItem(x, rect.top(), x, rect.bottom())
			line.setPen(pen)
			line.setZValue(9)
			self.scene.addItem(line)
			self.grid_items.append(line)
		
		for i in range(1, rows):
			y = rect.top() + i * cell_height
			line = QGraphicsLineItem(rect.left(), y, rect.right(), y)
			line.setPen(pen)
			line.setZValue(9)
			self.scene.addItem(line)
			self.grid_items.append(line)

	def set_preview_rects(self, preview_rects, rows: int, cols: int):
		for item in self.preview_items:
			self.scene.removeItem(item)
		self.preview_items = []
		
		if not preview_rects:
			return
		
		for i, rect in enumerate(preview_rects):
			row = i // cols
			col = i % cols
			if (row + col) % 2 == 0:
				color = QColor(255, 0, 0, 180)
			else:
				color = QColor(0, 255, 0, 180)
			pen = QPen(color, 2)
			item = QGraphicsRectItem(rect)
			item.setPen(pen)
			item.setBrush(Qt.BrushStyle.NoBrush)
			item.setZValue(8)
			self.scene.addItem(item)
			self.preview_items.append(item)

	def get_adjustment_type(self, pos: QPointF) -> int:
		"""根据鼠标位置返回调整类型 (场景坐标)"""
		# 优先检查手柄
		item = self.scene.itemAt(pos, self.transform())
		if item is not None:
			handle_type = item.data(0)
			if handle_type is not None:
				return int(handle_type)
		
		rect = self.selection_rect
		if rect.isNull() or rect.isEmpty():
			return self.NO_ADJUST
		
		scale = self.transform().m11()
		tol = self.HANDLE_SIZE / scale if scale != 0 else self.HANDLE_SIZE
		
		# 检查角部
		if math.hypot(pos.x() - rect.left(), pos.y() - rect.top()) < tol:
			return self.TOP_LEFT
		if math.hypot(pos.x() - rect.right(), pos.y() - rect.top()) < tol:
			return self.TOP_RIGHT
		if math.hypot(pos.x() - rect.left(), pos.y() - rect.bottom()) < tol:
			return self.BOTTOM_LEFT
		if math.hypot(pos.x() - rect.right(), pos.y() - rect.bottom()) < tol:
			return self.BOTTOM_RIGHT
		
		# 检查边缘
		if abs(pos.y() - rect.top()) <= tol and rect.left() <= pos.x() <= rect.right():
			return self.TOP
		if abs(pos.y() - rect.bottom()) <= tol and rect.left() <= pos.x() <= rect.right():
			return self.BOTTOM
		if abs(pos.x() - rect.left()) <= tol and rect.top() <= pos.y() <= rect.bottom():
			return self.LEFT
		if abs(pos.x() - rect.right()) <= tol and rect.top() <= pos.y() <= rect.bottom():
			return self.RIGHT
		
		# 检查内部（移动）
		if rect.contains(pos):
			return self.MOVE

		return self.NO_ADJUST

	def update_cursor(self, pos: QPointF):
		"""根据位置更新鼠标光标"""
		adjust_type = self.get_adjustment_type(pos)

		cursors = {
			self.TOP_LEFT: Qt.CursorShape.SizeFDiagCursor,
			self.TOP_RIGHT: Qt.CursorShape.SizeBDiagCursor,
			self.BOTTOM_LEFT: Qt.CursorShape.SizeBDiagCursor,
			self.BOTTOM_RIGHT: Qt.CursorShape.SizeFDiagCursor,
			self.TOP: Qt.CursorShape.SizeVerCursor,
			self.BOTTOM: Qt.CursorShape.SizeVerCursor,
			self.LEFT: Qt.CursorShape.SizeHorCursor,
			self.RIGHT: Qt.CursorShape.SizeHorCursor,
			self.MOVE: Qt.CursorShape.SizeAllCursor,
		}

		if adjust_type in cursors:
			self.setCursor(cursors[adjust_type])
		else:
			self.setCursor(Qt.CursorShape.OpenHandCursor)

	def mousePressEvent(self, event: QMouseEvent):
		if event.button() == Qt.LeftButton:
			pos_scene = self.mapToScene(event.position().toPoint())
			drag_type = self.get_adjustment_type(pos_scene)
			
			if drag_type != self.NO_ADJUST:
				# Dragging selection box
				self.is_dragging = True
				self.drag_start_pos = pos_scene
				self.original_rect = QRectF(self.selection_rect)
				self.drag_type = drag_type
				self.setFocus()
			else:
				# Dragging outside selection box - start panning
				self.is_panning = True
				self.pan_start_pos = event.position().toPoint()
				self.setCursor(Qt.CursorShape.ClosedHandCursor)

	def mouseMoveEvent(self, event: QMouseEvent):
		if self.is_panning:
			old_scene = self.mapToScene(self.pan_start_pos)
			new_scene = self.mapToScene(event.position().toPoint())
			delta = new_scene - old_scene
			self.translate(delta.x(), delta.y())
			self.pan_start_pos = event.position().toPoint()
			self.setCursor(Qt.CursorShape.ClosedHandCursor)
			return
		
		pos_scene = self.mapToScene(event.position().toPoint())
		if not self.is_dragging:
			self.update_cursor(pos_scene)
			return
	
		delta = pos_scene - self.drag_start_pos
		new_rect = QRectF(self.original_rect)

		# 根据调整类型更新选区
		if self.drag_type == self.MOVE:
			new_rect.translate(delta)
		elif self.drag_type == self.TOP_LEFT:
			new_rect.setTopLeft(self.original_rect.topLeft() + delta)
		elif self.drag_type == self.TOP_RIGHT:
			new_rect.setTopRight(self.original_rect.topRight() + delta)
		elif self.drag_type == self.BOTTOM_LEFT:
			new_rect.setBottomLeft(self.original_rect.bottomLeft() + delta)
		elif self.drag_type == self.BOTTOM_RIGHT:
			new_rect.setBottomRight(self.original_rect.bottomRight() + delta)
		elif self.drag_type == self.TOP:
			new_rect.setTop(self.original_rect.top() + delta.y())
		elif self.drag_type == self.BOTTOM:
			new_rect.setBottom(self.original_rect.bottom() + delta.y())
		elif self.drag_type == self.LEFT:
			new_rect.setLeft(self.original_rect.left() + delta.x())
		elif self.drag_type == self.RIGHT:
			new_rect.setRight(self.original_rect.right() + delta.x())

		# 确保选区在有效范围内
		if new_rect.width() > 10 and new_rect.height() > 10:
			self.selection_rect = new_rect
			self._update_selection_items()
			self.update_grid_items(self._parent.config.grid_rows, self._parent.config.grid_cols)
			# 更新归一化选区到配置
			if self._parent and self._parent.pixmap:
				self._parent.config.selection_x_normalized = self.selection_rect.x() / self._parent.pixmap.width()
				self._parent.config.selection_y_normalized = self.selection_rect.y() / self._parent.pixmap.height()
				self._parent.config.selection_w_normalized = self.selection_rect.width() / self._parent.pixmap.width()
				self._parent.config.selection_h_normalized = self.selection_rect.height() / self._parent.pixmap.height()

	def mouseReleaseEvent(self, event: QMouseEvent):
		if event.button() == Qt.LeftButton:
			if self.is_panning:
				self.is_panning = False
				self.setCursor(Qt.CursorShape.ArrowCursor)
			elif self.is_dragging:
				self.is_dragging = False
				self._parent.update_preview()
	
	def wheelEvent(self, event: QWheelEvent):
		"""处理滚轮事件进行缩放（以光标为中心）"""
		if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
			angle_delta = event.angleDelta().y()
			if angle_delta == 0:
				return
			
			old_pos = self.mapToScene(event.position().toPoint())
			factor = 1.1 if angle_delta > 0 else 1 / 1.1
			new_scale = self._parent.config.image_scale * factor
			new_scale = max(0.1, min(10.0, new_scale))
			factor = new_scale / self._parent.config.image_scale
			self._parent.config.image_scale = new_scale
			
			self.scale(factor, factor)
			new_pos = self.mapToScene(event.position().toPoint())
			delta = new_pos - old_pos
			self.translate(delta.x(), delta.y())
			
			event.accept()
		else:
			super().wheelEvent(event)

class ImageGridSplitter(QMainWindow):
	"""主窗口类"""

	def __init__(self):
		super().__init__()
		self.config = AppConfig.load()
		self.current_image_path = None
		self.pixmap = None
		self.scaled_pixmap = None
		self.image_rect = QRect()
		self.preview_rects = []

		self.init_ui()
		self.setAcceptDrops(True)
		self.resize(self.config.window_width, self.config.window_height)

	def init_ui(self):
		"""初始化用户界面"""
		self.setWindowTitle('图片网格分割工具')
		self.setWindowIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView))

		# 创建中央部件
		central_widget = QWidget()
		self.setCentralWidget(central_widget)

		# 主布局
		main_layout = QVBoxLayout(central_widget)

		# 工具栏
		toolbar = self.create_toolbar()
		toolbar.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
		main_layout.addWidget(toolbar)

		# 图片显示区域
		self.image_label = DraggableSelectionBox(self)
		self.image_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
		self.image_label.setMinimumSize(400, 300)
		main_layout.addWidget(self.image_label, 1)

		# 控制面板
		control_panel = self.create_control_panel()
		control_panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
		main_layout.addWidget(control_panel)

		# 状态栏
		self.statusBar().showMessage('拖放图片文件到窗口开始使用')

	def create_toolbar(self) -> QWidget:
		"""创建工具栏"""
		toolbar = QWidget()
		toolbar.setMaximumHeight(50)
		layout = QHBoxLayout(toolbar)
		layout.setContentsMargins(10, 5, 10, 5)

		# 打开按钮
		self.open_btn = QPushButton('打开图片')
		self.open_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
		self.open_btn.clicked.connect(self.open_image)
		layout.addWidget(self.open_btn)

		# 分割按钮
		self.split_btn = QPushButton('分割图片')
		self.split_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
		self.split_btn.clicked.connect(self.split_image)
		layout.addWidget(self.split_btn)

		# PDF导出按钮
		self.pdf_btn = QPushButton('导出PDF')
		self.pdf_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
		self.pdf_btn.clicked.connect(self.export_pdf)
		layout.addWidget(self.pdf_btn)

		layout.addStretch()

		# 裁剪边框复选框
		self.cut_border_checkbox = QCheckBox('裁剪边框')
		self.cut_border_checkbox.setChecked(self.config.cut_border)
		self.cut_border_checkbox.stateChanged.connect(self.toggle_cut_border)
		layout.addWidget(self.cut_border_checkbox)

		# 预览模式复选框
		self.preview_checkbox = QCheckBox('预览模式')
		self.preview_checkbox.setChecked(self.config.preview_mode)
		self.preview_checkbox.stateChanged.connect(self.toggle_preview)
		layout.addWidget(self.preview_checkbox)

		# 显示信息
		self.info_label = QLabel('')
		layout.addWidget(self.info_label)

		return toolbar

	def create_control_panel(self) -> QWidget:
		"""创建控制面板"""
		panel = QWidget()
		panel.setMaximumHeight(100)
		layout = QHBoxLayout(panel)
		layout.setContentsMargins(15, 10, 15, 10)

		rows_cols_layout = QFormLayout()

		# 行数控制

		self.rows_spin = QSpinBox()
		self.rows_spin.setMinimum(1)
		self.rows_spin.setValue(self.config.grid_rows)
		self.rows_spin.setFixedWidth(80)
		self.rows_spin.valueChanged.connect(self.update_grid)

		rows_cols_layout.addRow('行数', self.rows_spin)

		# 列数控制

		self.cols_spin = QSpinBox()
		self.cols_spin.setMinimum(1)
		self.cols_spin.setValue(self.config.grid_cols)
		self.cols_spin.setFixedWidth(80)
		self.cols_spin.valueChanged.connect(self.update_grid)

		rows_cols_layout.addRow('列数', self.cols_spin)

		layout.addLayout(rows_cols_layout)
		# layout.addLayout(quick_btn_layout)
		layout.addStretch()

		return panel

	def update_grid(self):
		"""更新网格设置"""
		self.config.grid_rows = self.rows_spin.value()
		self.config.grid_cols = self.cols_spin.value()

		# Update preview if in preview mode
		if self.config.preview_mode and self.pixmap:
			self.update_preview()
		
		self.image_label.update_grid_items(self.config.grid_rows, self.config.grid_cols)

		# 更新信息显示
		if self.pixmap:
			self.update_info()

	def update_info(self):
		"""更新信息显示"""
		if not self.pixmap:
			return
	
		total = self.config.grid_rows * self.config.grid_cols
		if self.current_image_path:
			info = f"{self.current_image_path} | 网格: {self.config.grid_rows}×{self.config.grid_cols} = {total}张图片"
		else:
			# For clipboard images without a file path
			info = f"剪贴板图片 ({self.pixmap.width()}×{self.pixmap.height()}) | 网格: {self.config.grid_rows}×{self.config.grid_cols} = {total}张图片"
		self.info_label.setText(info)

	def update_preview(self):
		"""更新预览边界框"""
		if not (self.config.preview_mode and self.pixmap):
			self.preview_rects = []
			self.image_label.set_preview_rects([], self.config.grid_rows, self.config.grid_cols)
			return

		self.preview_rects = []

		# Calculate grid cells in display coordinates
		label_rect = self.image_label.selection_rect
		rows = self.config.grid_rows
		cols = self.config.grid_cols

		cell_width = label_rect.width() / cols
		cell_height = label_rect.height() / rows

		for row in range(rows):
			for col in range(cols):
				# Calculate cell position in display coordinates
				x_display = label_rect.left() + col * cell_width
				y_display = label_rect.top() + row * cell_height
				w_display = cell_width if col < cols - 1 else (label_rect.right() - x_display)
				h_display = cell_height if row < rows - 1 else (label_rect.bottom() - y_display)
		
				if self.config.cut_border:
					# Convert to original image coordinates (scene == image)
					x_orig = round(x_display)
					y_orig = round(y_display)
					w_orig = round(w_display)
					h_orig = round(h_display)
			
					# Get image data
					img_data = self.pixmap.toImage()
					# Convert to numpy array
					width = img_data.width()
					height = img_data.height()
					ptr = img_data.bits()
					if ptr is None:
						continue
					arr = np.array(ptr).reshape(height, width, 4)
			
					# Crop to cell
					x2 = min(x_orig + w_orig, width)
					y2 = min(y_orig + h_orig, height)
					cell_arr = arr[y_orig:y2, x_orig:x2, :3]  # RGB only
			
					try:
						# Detect border
						top_crop, bottom_crop, left_crop, right_crop = self.detect_border_with_otsu(cell_arr)
			
						# Adjust display coordinates
						x_display += left_crop
						y_display += top_crop
						w_display = (right_crop - left_crop)
						h_display = (bottom_crop - top_crop)
					except:
						pass  # If border detection fails, use original coordinates
		
				preview_rect = QRectF(x_display, y_display, w_display, h_display)
				self.preview_rects.append(preview_rect)

		self.image_label.set_preview_rects(self.preview_rects, rows, cols)

	def open_image(self):
		"""打开图片文件"""
		file_path, _ = QFileDialog.getOpenFileName(
			self, '选择图片', '',
			'图片文件 (*.png *.jpg *.jpeg *.bmp *.gif *.tiff)'
		)

		if file_path:
			self.load_image(file_path)

	def load_image(self, file_path: str):
		"""加载图片"""
		self.pixmap = QPixmap(file_path)
		if self.pixmap.isNull():
			QMessageBox.warning(self, '错误', '无法加载图片文件！')
			return
	
		self.current_image_path = file_path
		self.scale_image()
		self.update_info()
		self.statusBar().showMessage(f'已加载: {os.path.basename(file_path)}')

	def paste_image_from_clipboard(self):
		"""从剪贴板粘贴图片"""
		clipboard = QApplication.clipboard()
		mime_data = clipboard.mimeData()

		# Check if clipboard has image data
		if mime_data.hasImage():
			pixmap = clipboard.pixmap()
			if not pixmap.isNull():
				# Load directly from pixmap without temporary file
				self.load_image_from_pixmap(pixmap)
				return

		# Check if clipboard has file URLs (like from file manager)
		if mime_data.hasUrls():
			urls = mime_data.urls()
			for url in urls:
				file_path = url.toLocalFile()
				if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff')):
					self.load_image(file_path)
					return

		# If no image found
		self.statusBar().showMessage('剪贴板中没有图片或图片文件路径')

	def load_image_from_pixmap(self, pixmap: QPixmap):
		"""从QPixmap直接加载图片（用于剪贴板粘贴）"""
		if pixmap.isNull():
			QMessageBox.warning(self, '错误', '无法加载图片！')
			return
	
		self.pixmap = pixmap
		self.current_image_path = None  # No file path for clipboard images
		self.scale_image()
		self.update_info()
		self.statusBar().showMessage('已从剪贴板加载图片')

	def scale_image(self):
		"""缩放图片以适应显示区域"""
		if not self.pixmap:
			return
		self.image_label.set_pixmap(self.pixmap, apply_fit=True)
		
		img_rect = QRectF(0, 0, self.pixmap.width(), self.pixmap.height())
		selection_rect = QRectF(
			self.pixmap.width() * self.config.selection_x_normalized,
			self.pixmap.height() * self.config.selection_y_normalized,
			self.pixmap.width() * self.config.selection_w_normalized,
			self.pixmap.height() * self.config.selection_h_normalized
		)
		selection_rect = selection_rect.intersected(img_rect)
		if selection_rect.width() < 10 or selection_rect.height() < 10:
			selection_rect = QRectF(
				self.pixmap.width() * 0.1,
				self.pixmap.height() * 0.1,
				self.pixmap.width() * 0.8,
				self.pixmap.height() * 0.8
			)
		self.image_label.set_selection_rect(selection_rect, self.config.grid_rows, self.config.grid_cols)
		
		self.update_preview()

	def toggle_cut_border(self, state):
		"""切换裁剪边框选项"""
		self.config.cut_border = (state == Qt.CheckState.Checked.value)
		self.update_preview()

	def toggle_preview(self, state):
		"""切换预览模式"""
		self.config.preview_mode = (state == Qt.CheckState.Checked.value)
		self.update_preview()

	def detect_border_with_otsu(self, img_array):
		"""使用Otsu方法检测并返回边框裁剪区域"""
		# Convert to grayscale if needed
		if len(img_array.shape) == 3:
			gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
		else:
			gray = img_array

		# Apply Otsu's thresholding
		_, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

		# Sum along axes to find borders
		vertical_sum = np.sum(binary, axis=0)
		horizontal_sum = np.sum(binary, axis=1)

		# Get border values (should be same at both ends)
		top_val = binary[0, 0]
		bottom_val = binary[-1, -1]
		left_val = binary[0, 0]
		right_val = binary[-1, -1]

		# Find crop boundaries
		# From top
		top_crop = 0
		for i in range(len(horizontal_sum)):
			if np.mean(binary[i, :]) != top_val:
				top_crop = i
				break

		# From bottom
		bottom_crop = len(horizontal_sum)
		for i in range(len(horizontal_sum) - 1, -1, -1):
			if np.mean(binary[i, :]) != bottom_val:
				bottom_crop = i + 1
				break

		# From left
		left_crop = 0
		for i in range(len(vertical_sum)):
			if np.mean(binary[:, i]) != left_val:
				left_crop = i
				break

		# From right
		right_crop = len(vertical_sum)
		for i in range(len(vertical_sum) - 1, -1, -1):
			if np.mean(binary[:, i]) != right_val:
				right_crop = i + 1
				break

		return top_crop, bottom_crop, left_crop, right_crop

	def get_split_images(self):
		"""获取分割后的图片数组列表"""
		if not self.pixmap:
			return []

		images = []

		# Calculate selection in original image coordinates
		label_rect = self.image_label.selection_rect
		img_rect = QRect(
			round(label_rect.left()),
			round(label_rect.top()),
			round(label_rect.width()),
			round(label_rect.height())
		)

		img_rect = img_rect.intersected(QRect(0, 0, self.pixmap.width(), self.pixmap.height()))

		if img_rect.isEmpty():
			return []

		# Convert QPixmap to numpy array
		img_data = self.pixmap.toImage()
		width = img_data.width()
		height = img_data.height()
		ptr = img_data.bits()
		if ptr is None:
			return []
		arr = np.array(ptr).reshape(height, width, 4)[:, :, :3][..., ::-1]  # RGB only, but PySide6 stores RGB as BGR

		# Calculate cell sizes
		cell_width = img_rect.width() // self.config.grid_cols
		cell_height = img_rect.height() // self.config.grid_rows

		for row in range(self.config.grid_rows):
			for col in range(self.config.grid_cols):
				x = img_rect.left() + col * cell_width
				y = img_rect.top() + row * cell_height
		
				# Handle last column/row
				if col == self.config.grid_cols - 1:
					w = img_rect.right() - x
				else:
					w = cell_width
		
				if row == self.config.grid_rows - 1:
					h = img_rect.bottom() - y
				else:
					h = cell_height
		
				# Crop cell
				cell_arr = arr[y:y+h, x:x+w].copy()
		
				# Cut border if needed
				if self.config.cut_border:
					try:
						top_crop, bottom_crop, left_crop, right_crop = self.detect_border_with_otsu(cell_arr)
						cell_arr = cell_arr[top_crop:bottom_crop, left_crop:right_crop]
					except:
						pass  # Keep original if border detection fails
		
				images.append(cell_arr)

		return images

	def split_image(self):
		"""分割图片"""
		if not self.pixmap:
			QMessageBox.warning(self, '错误', '请先加载图片！')
			return

		# Determine default directory and base name
		if self.current_image_path:
			default_dir = os.path.dirname(self.current_image_path)
			base_name = os.path.splitext(os.path.basename(self.current_image_path))[0]
		else:
			# For clipboard images
			default_dir = os.path.expanduser('~')
			base_name = 'clipboard_image'

		# 选择保存目录
		save_dir = QFileDialog.getExistingDirectory(
			self, '选择保存目录',
			default_dir
		)

		if not save_dir:
			return

		# Get split images using numpy
		images = self.get_split_images()

		if not images:
			QMessageBox.warning(self, '错误', '无法获取分割图片！')
			return

		# Save images using PIL
		saved_count = 0
		idx = 0
		for row in range(self.config.grid_rows):
			for col in range(self.config.grid_cols):
				if idx >= len(images):
					break
		
				# Convert numpy array to PIL Image
				pil_img = Image.fromarray(images[idx])
		
				# Save image
				save_path = os.path.join(
					save_dir,
					f'{base_name}_r{row+1}c{col+1}.png'
				)
		
				try:
					pil_img.save(save_path, 'PNG')
					saved_count += 1
				except:
					pass
		
				idx += 1

		QMessageBox.information(
			self, '完成',
			f'成功分割并保存了 {saved_count} 张图片到:\n{save_dir}'
		)

	def export_pdf(self):
		"""导出PDF"""
		if not self.pixmap:
			QMessageBox.warning(self, '错误', '请先加载图片！')
			return

		# Create dialog for PDF settings
		dialog = QDialog(self)
		dialog.setWindowTitle('PDF导出设置')
		dialog_layout = QVBoxLayout(dialog)

		# Page size selection
		size_group = QGroupBox('页面尺寸')
		size_layout = QHBoxLayout()

		size_combo = QComboBox()
		size_combo.addItems(['A4', 'Letter', '16:9', '4:3', '自定义'])
		size_combo.setCurrentText(self.config.pdf_preset)
		size_layout.addWidget(QLabel('预设:'))
		size_layout.addWidget(size_combo)

		width_spin = QDoubleSpinBox()
		width_spin.setRange(5, 100)
		width_spin.setValue(self.config.pdf_width_spin)
		width_spin.setSuffix(' cm')
		width_spin.setDecimals(1)

		height_spin = QDoubleSpinBox()
		height_spin.setRange(5, 100)
		height_spin.setValue(self.config.pdf_height_spin)
		height_spin.setSuffix(' cm')
		height_spin.setDecimals(1)

		def update_size():
			preset = size_combo.currentText()
			if preset == 'A4':
				width_spin.setValue(A4[0] / cm)
				height_spin.setValue(A4[1] / cm)
				width_spin.setEnabled(False)
				height_spin.setEnabled(False)
			elif preset == 'Letter':
				width_spin.setValue(letter[0] / cm)
				height_spin.setValue(letter[1] / cm)
				width_spin.setEnabled(False)
				height_spin.setEnabled(False)
			elif preset == '16:9':
				width_spin.setValue(33.867)
				height_spin.setValue(19.05)
				width_spin.setEnabled(False)
				height_spin.setEnabled(False)
			elif preset == '4:3':
				width_spin.setValue(25.4)
				height_spin.setValue(19.05)
				width_spin.setEnabled(False)
				height_spin.setEnabled(False)
			else:
				width_spin.setEnabled(True)
				height_spin.setEnabled(True)

		size_combo.currentTextChanged.connect(update_size)
		update_size()

		size_layout.addWidget(QLabel('宽度:'))
		size_layout.addWidget(width_spin)
		size_layout.addWidget(QLabel('高度:'))
		size_layout.addWidget(height_spin)

		size_group.setLayout(size_layout)
		dialog_layout.addWidget(size_group)

		# Buttons
		button_box = QHBoxLayout()
		ok_btn = QPushButton('确定')
		cancel_btn = QPushButton('取消')
		ok_btn.clicked.connect(dialog.accept)
		cancel_btn.clicked.connect(dialog.reject)
		button_box.addWidget(ok_btn)
		button_box.addWidget(cancel_btn)
		dialog_layout.addLayout(button_box)

		if dialog.exec() != QDialog.DialogCode.Accepted:
			return

		self.config.pdf_preset = size_combo.currentText()
		self.config.pdf_width_spin = width_spin.value()
		self.config.pdf_height_spin = height_spin.value()

		# Determine default save path
		if self.current_image_path:
			default_path = os.path.splitext(self.current_image_path)[0] + '.pdf'
		else:
			default_path = os.path.join(os.path.expanduser('~'), 'clipboard_image.pdf')

		# Get PDF file path
		save_path, _ = QFileDialog.getSaveFileName(
			self, '保存PDF',
			default_path,
			'PDF文件 (*.pdf)'
		)

		if not save_path:
			return

		# Get page size in points (1 cm = 28.3465 points)
		page_width = self.config.pdf_width_spin * cm
		page_height = self.config.pdf_height_spin * cm

		# Create PDF
		try:
			c = canvas.Canvas(save_path, pagesize=(page_width, page_height))
	
			# Get all split images
			images = self.get_split_images()
	
			if not images:
				QMessageBox.warning(self, '错误', '无法获取分割图片！')
				return
	
			for img_array in images:
				# Convert numpy array to PIL Image
				pil_img = Image.fromarray(img_array)
		
				# Get median border value for padding
				border = np.median(np.concatenate([img_array[0, :], img_array[-1, :], img_array[:, 0], img_array[:, -1]]), axis=0)
				median_color = tuple(int(i) for i in border)
		
				# Set background color so that image will not distort by resizing
				c.setFillColorRGB(*[i/255 for i in median_color])
				c.rect(0,0,page_width,page_height,fill=1)
		
				# Draw on PDF
				img_buffer = io.BytesIO()
				pil_img.save(img_buffer, format='PNG')
				img_buffer.seek(0)
		
				c.drawImage(ImageReader(img_buffer), 0, 0, page_width, page_height, preserveAspectRatio=True) # pdf is svg-like so it will automatically scale
				c.showPage()
	
			c.save()
			QMessageBox.information(self, '完成', f'PDF已保存到:\n{save_path}')
	
		except Exception as e:
			QMessageBox.critical(self, '错误', f'PDF导出失败:\n{str(e)}')

	# 拖放功能
	def dragEnterEvent(self, event: QDragEnterEvent):
		if event.mimeData().hasUrls():
			event.acceptProposedAction()

	def dropEvent(self, event: QDropEvent):
		urls = event.mimeData().urls()
		if urls:
			file_path = urls[0].toLocalFile()
			if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff')):
				self.load_image(file_path)
				event.acceptProposedAction()

	def keyPressEvent(self, event: QKeyEvent):
		"""处理键盘事件，包括Ctrl+V粘贴图片，Ctrl+0重置缩放和平移"""
		# Check if Ctrl+V is pressed and focus is not on spinbox
		if event.key() == Qt.Key.Key_V and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
			# Check if focus is on spinbox - if so, let it handle paste
			focused_widget = QApplication.focusWidget()
			if isinstance(focused_widget, QSpinBox):
				super().keyPressEvent(event)
				return
	
			# Try to paste image from clipboard
			self.paste_image_from_clipboard()
		else:
			super().keyPressEvent(event)

		if event.key() == Qt.Key.Key_0 and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
			# Reset image translation and scale
			# self.config.image_translation_x_normalized = 0.0
			# self.config.image_translation_y_normalized = 0.0
			self.config.image_scale = 1.0
			self.scale_image()

	# 窗口事件
	def resizeEvent(self, event: QResizeEvent):
		"""窗口大小改变时的事件"""
		super().resizeEvent(event)
		# Scale image when window is resized
		if self.pixmap:
			self.scale_image()

		# 保存窗口大小
		self.config.window_width = self.width()
		self.config.window_height = self.height()

	def closeEvent(self, event: QCloseEvent):
		"""关闭窗口时保存配置"""
		self.config.save()
		event.accept()

def main():
	"""主函数"""
	app = QApplication(sys.argv)
	app.setStyle('Fusion')  # 使用Fusion风格，更美观

	# 设置应用程序样式
	app.setStyleSheet("""
		QMainWindow {
			background-color: #323232;
		}
		QGroupBox {
			font-weight: bold;
			border: 1px solid #555;
			border-radius: 4px;
			margin-top: 10px;
			padding-top: 10px;
		}
		QGroupBox::title {
			subcontrol-origin: margin;
			left: 10px;
			padding: 0 5px 0 5px;
		}
		QPushButton {
			background-color: #505050;
			border: 1px solid #555;
			border-radius: 3px;
			padding: 5px 15px;
			min-width: 80px;
		}
		QPushButton:hover {
			background-color: #606060;
			border: 1px solid #666;
		}
		QPushButton:pressed {
			background-color: #404040;
		}
		QSpinBox {
			padding: 3px;
			border: 1px solid #555;
			border-radius: 3px;
			background-color: #404040;
		}
		QToolButton {
			background-color: #505050;
			border: 1px solid #555;
			border-radius: 2px;
		}
		QToolButton:hover {
			background-color: #606060;
		}
		QStatusBar {
			background-color: #2b2b2b;
			color: #aaa;
		}
	""")

	window = ImageGridSplitter()
	window.show()

	sys.exit(app.exec())

if __name__ == '__main__':
	main()