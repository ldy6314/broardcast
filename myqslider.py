

from PyQt5.QtWidgets import QSlider


class MySlider(QSlider):
    def __init__(self, parent=None):
        super(MySlider, self).__init__(parent)

    def mousePressEvent(self, event):
        p = event.pos().x() / self.width()
        val = p * (self.maximum() - self.minimum()) + self.minimum()
        self.setValue(val)
        super().mousePressEvent(event)
