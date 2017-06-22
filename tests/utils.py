
class FakeCell:
    coordinate = "A1"

    def __init__(self, value):
        self.value = value

    @classmethod
    def create(cls, values):
        return tuple(cls(value) for value in values)


def FakeCells(*values):
    return tuple(FakeCell(value) for value in values)
