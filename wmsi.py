import math
import sys

from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QColor, QBrush, QPen, QPainterPath
from PySide6.QtWidgets import (
    QApplication,
    QGraphicsPathItem,
    QGraphicsScene,
    QGraphicsView,
    QLabel,
    QMainWindow,
    QVBoxLayout,
    QWidget,
)

# -----------------------------
# Modelo
# -----------------------------

COLORS = {
    1: QColor("#00d000"),   # normal
    2: QColor("#ffd000"),   # hipocinético
    3: QColor("#0099ff"),   # acinético
    4: QColor("#ff3030"),   # discinético
}


class HeartModel(QObject):

    changed = Signal()

    def __init__(self):
        super().__init__()

        self.scores = {i: 1 for i in range(1, 18)}

    def next_score(self, segment):

        s = self.scores[segment] + 1

        if s > 4:
            s = 1

        self.scores[segment] = s

        self.changed.emit()

    def score(self, segment):
        return self.scores[segment]

    def wmsi(self):
        return sum(self.scores.values()) / 17


# -----------------------------
# Segmento gráfico
# -----------------------------

class SegmentItem(QGraphicsPathItem):

    def __init__(self, segment, path, model):
        super().__init__(path)

        self.segment = segment
        self.model = model

        self.setPen(QPen(Qt.black, 2))
        self.update_color()

        model.changed.connect(self.update_color)

    def update_color(self):
        self.setBrush(QBrush(COLORS[self.model.score(self.segment)]))

    def mousePressEvent(self, event):
        self.model.next_score(self.segment)


# -----------------------------
# Bullseye
# -----------------------------

class Bullseye(QGraphicsView):

    def __init__(self, model):
        super().__init__()

        self.model = model

        scene = QGraphicsScene()
        self.setScene(scene)

        self.setRenderHints(self.renderHints())

        self.build(scene)

        self.setFixedSize(420,420)

    def sector(self, r1, r2, a1, a2):

        path = QPainterPath()

        x = 0
        y = 0

        # arco externo

        rect2 = (-r2,-r2,r2*2,r2*2)
        rect1 = (-r1,-r1,r1*2,r1*2)

        path.arcMoveTo(*rect2,a1)
        path.arcTo(*rect2,a1,a2-a1)

        path.lineTo(
            math.cos(math.radians(a2))*r1,
            -math.sin(math.radians(a2))*r1
        )

        path.arcTo(*rect1,a2,a1-a2)

        path.closeSubpath()

        return path

    def build(self,scene):

        segment = 1

        rings = [
            (120,160,6),
            (80,120,6),
            (40,80,4)
        ]

        for r1,r2,n in rings:

            step = 360/n

            for i in range(n):

                a1 = i*step
                a2 = (i+1)*step

                item = SegmentItem(
                    segment,
                    self.sector(r1,r2,a1,a2),
                    self.model
                )

                scene.addItem(item)

                segment+=1

        # centro

        center = QPainterPath()
        center.addEllipse(-40,-40,80,80)

        item = SegmentItem(17,center,self.model)
        scene.addItem(item)

        scene.setSceneRect(-180,-180,360,360)


# -----------------------------
# Janela
# -----------------------------

class MainWindow(QMainWindow):

    def __init__(self):

        super().__init__()

        self.model = HeartModel()

        self.setWindowTitle("Wall Motion")

        widget = QWidget()
        self.setCentralWidget(widget)

        layout = QVBoxLayout(widget)

        # -------------------
        # Bullseye
        # -------------------
        self.bull = Bullseye(self.model)
        layout.addWidget(self.bull)

        # -------------------
        # WMSI label
        # -------------------
        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("""
            font-size:18px;
            font-weight:bold;
            color:white;
        """)

        layout.addWidget(self.label)

        # -------------------
        # LEGENDA
        # -------------------
        legend = self.create_legend()
        layout.addWidget(legend)

        self.model.changed.connect(self.update_label)
        self.update_label()

    def update_label(self):

        self.label.setText(
            f"WMSI = {self.model.wmsi():.2f}"
        )

    def create_legend(self):

        w = QWidget()
        layout = QVBoxLayout(w)

        title = QLabel("Legenda")
        title.setStyleSheet("color:white; font-weight:bold;")
        layout.addWidget(title)

        items = [
            ("Normal (1)", COLORS[1]),
            ("Hipocinético (2)", COLORS[2]),
            ("Acinético (3)", COLORS[3]),
            ("Discinético (4)", COLORS[4]),
        ]

        for text, color in items:

            row = QWidget()
            row_layout = QVBoxLayout(row)

            label = QLabel(text)
            label.setStyleSheet(f"""
                color:white;
            """)

            color_box = QLabel()
            color_box.setFixedSize(18, 18)
            color_box.setStyleSheet(f"""
                background:{color.name()};
                border:1px solid black;
            """)

            line = QVBoxLayout()
            line_widget = QWidget()
            line_widget.setStyleSheet("margin:2px;")

            hl = QVBoxLayout()
            line_widget.setLayout(hl)

            hl.addWidget(color_box)
            hl.addWidget(label)

            layout.addWidget(line_widget)

        return w

# -----------------------------

app = QApplication(sys.argv)

app.setStyleSheet("""

QMainWindow{
background:#2d2d30;
}

QLabel{
color:white;
}

QGraphicsView{
background:#202020;
border:none;
}

""")

window = MainWindow()
window.resize(500,500)
window.show()

app.exec()
