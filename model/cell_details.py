from model.coordinate import Coordinate


class CellDetails:
    def __init__(self, coordinate: Coordinate, cell_format: str, value: any) -> None:
        self.coordinate: Coordinate = coordinate
        self.cell_format: str = cell_format
        self.value: any = value
