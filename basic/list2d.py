"""
    This module has two classes, "Matrix" and "Table".
    The form of the two classes is two-dimensional list.
"""

from typing import TypeVar, Generic, List
from itertools import product, combinations

T = TypeVar('T')
VHT = TypeVar('VHT')
HHT = TypeVar('HHT')


class Matrix(Generic[T]):
    """ Simple matrix.

        Attributes:
            _matrix: two-dimensional list that has data
            _n_row: the number of rows
            _n_col: the number of columns
    """

    def __init__(self):
        self._matrix: List[List[T]] = []
        self._n_row = 0
        self._n_col = 0

    # Getters
    @property
    def n_row(self):
        return self._n_row

    @property
    def n_col(self):
        return self._n_col

    def contents(self) -> List[List[T]]:
        return self._matrix


    def insert(self, value: T, row: int, col: int) -> None:
        """ Inserts value in the specific position.

        If the matrix is smaller than "row" or "col", it will expand.
        The value of empty cells is "None".
        :param value: data which will be inserted
        :param row: row position
        :param col: column position
        """
        if (self._n_row > row and self._n_col > col) is False:
            for i in range(row - self._n_row + 1):
                self._matrix.append([None] * self._n_col)
            if self._n_row <= row:
                self._n_row = row + 1

            for r in range(self._n_row):
                if col >= self._n_col:
                    self._matrix[r] = self._matrix[r] + [None] * (col - self._n_col + 1)
            if self._n_col <= col:
                self._n_col = col + 1

        self._matrix[row][col] = value

    def insert_row(self, row: int = None) -> None:
        """ Inserts an empty row in the specific position.

        The value of empty cells is "None".
        :param row: inserted row position. If it is "None", a row will be placed at the last.
        """
        if row is None:
            row = self._n_row
        self._matrix.insert(row, [None] * self._n_col)
        self._n_row += 1

    def insert_col(self, col: int = None) -> None:
        """ Inserts an empty column in the specific position.

        The value of empty cells is "None".
        :param col: inserted column position. If it is "None", a column will be placed at the last.
        """
        if col is None:
            col = self._n_col
        for r in self._matrix:
            r.insert(col, None)
        self._n_col += 1

    def get(self, row: int, col: int) -> T:
        """ Returns data in the specific position.

        :param row: row position
        :param col: column position
        :return: data in the specific position.
        """
        return self._matrix[row][col]

    def get_row(self, row: int) -> List[T]:
        """ Returns a list of data in the specific row.

        :param row: row position
        :return: a list of data in the specific row.
        """
        return self._matrix[row]

    def get_col(self, col: int) -> List[T]:
        """ Returns a list of data in the specific column.

        :param col: column position
        :return: a list of data in the specific column.
        """
        result = []
        for r in self._matrix:
            result.append(r[col])
        return result

    def remove_row(self, row: int = None):
        """ Removes a row.

        :param row: a row which will be removed.
        """
        if row is None:
            row = self._n_row
        del self._matrix[row]
        self._n_row -= 1

    def remove_col(self, col: int = None):
        """ Removes a column.

        :param col: a column which will be removed.
        """
        if col is None:
            col = self._n_col
        for r in self._matrix:
            del r[col]
        self._n_col -= 1

    def __str__(self):
        """ Print the object as a matrix """
        col_space = [0] * self._n_col
        for i, r in zip(range(self._n_col), self._matrix):
            for cell in r:
                if col_space[i] < len(cell.__str__()):
                    col_space[i] = len(cell.__str__())

        result = ""
        for r in self._matrix:
            for i, cell in enumerate(r):
                result += ("%" + str(col_space[i]) + "s    ") % cell
            result += "\n"
        return result

    def __eq__(self, other):
        if isinstance(other, Matrix):
            if self.n_row != other.n_row or self.n_col != other.n_col:
                return False

            for row, col in product(range(self._n_row), range(self._n_col)):
                if self.get(row, col) != other.get(row, col):
                    return False
            return True

        return super(Matrix, self).__eq__(other)

    def __getitem__(self, item: int):
        return self._matrix[item]


class Table(Matrix[T], Generic[T, VHT, HHT]):
    """ This class represents a table.

        Inherits "Matrix" class.

        Attributes:
            _header_v: a list of vertical (row) header
            _header_h: a list of horizontal (column) header
    """

    def __init__(self):
        super(Table, self).__init__()
        self._header_v: List[VHT] = []
        self._header_h: List[HHT] = []

    # Getters
    @property
    def header_v(self):
        return self._header_v

    @property
    def header_h(self):
        return self._header_h

    def get_with_header(self, row_h: VHT, col_h: HHT) -> T:
        """ Returns data corresponding to the specified row header and column header.

        :param row_h: row header of data which will be returned
        :param col_h: column header of data which will be returned
        :return: data corresponding to the specified row header and column header.
        """
        return self.get(self._header_v.index(row_h), self._header_h.index(col_h))

    def get_row_with_header(self, row_h: VHT) -> List[T]:
        """ Returns a list of row specified by row header.

        :param row_h: row header of data which will be returned
        :return: a list of row specified by row header.
        """
        return self.get_row(self._header_v.index(row_h))

    def get_col_with_header(self, col_h: HHT) -> List[T]:
        """ Returns a list of column specified by column header.

        :param col_h: column header of data which will be returned
        :return: a list of column specified by column header.
        """
        return self.get_col(self._header_h.index(col_h))

    def remove_row(self, row: int = None):
        """ Removes a row. Corresponding header also will be deleted.

        :param row: a row which will be removed.
        """
        super(Table, self).remove_row(row)
        del self._header_v[row]

    def remove_row_with_header(self, row_h: VHT):
        """ Removes a row by header. Corresponding header also will be deleted.

        :param row_h: a row which will be removed.
        """
        self.remove_row(self._header_v.index(row_h))

    def remove_col(self, col: int = None):
        """ Removes a column. Corresponding header also will be deleted.

        :param col: a column which will be removed.
        """
        super(Table, self).remove_col(col)
        del self._header_h[col]

    def remove_col_with_header(self, col_h: HHT):
        """ Removes a column by header. Corresponding header also will be deleted.

        :param col_h: a column which will be removed.
        """
        self.remove_col(self._header_h.index(col_h))

    def insert(self, value: T, row: int, col: int):
        """Inserts value in the specific position.

        If the matrix is smaller than "row" or "col", it will expand.
        If it expands, headers also expand, assigning empty headers "None".
        The value of empty cells is "None".

        :param value: data which will be inserted
        :param row: row position
        :param col: column position
        """
        if row >= len(self._header_v):
            self._header_v = self._header_v + [None] * (row - len(self._header_v) + 1)
        if col >= len(self._header_h):
            self._header_h = self._header_h + [None] * (col - len(self._header_h) + 1)
        super(Table, self).insert(value, row, col)

    def insert_with_header(self, value: T, row_h: VHT, col_h: HHT):
        """ Inserts value in the specific position specified by headers.

        :param value: data which will be inserted
        :param row_h: row header specifying position.
        :param col_h: column header specifying position.
        """
        self.insert(value, self._header_v.index(row_h), self._header_h.index(col_h))

    def insert_header_v(self, row_h: VHT, row: int = None):
        """ Inserts vertical header in the specific position.

        An empty row will be inserted where the header is inserted.

        :param row_h: vertical header which will be inserted.
        :param row: row position. If it is None, header will be inserted at the last.
        :raise Table.SameHeaderExistError: if there is the same header.
        :raise ValueError: if row_h is None.
        """
        if row_h is None:
            raise ValueError

        try:
            self._header_v.index(row_h)
            raise Table.SameHeaderExistError
        except ValueError:
            pass

        if row is None:
            row = self._n_row
        self.insert_row(row)
        self._header_v.insert(row, row_h)

    def insert_header_h(self, col_h: HHT, col: int = None):
        """ Inserts horizontal header in the specific position.

        An empty column will be inserted where the header is inserted.

        :param col_h: horizontal header which will be inserted.
        :param col: column position. If it is None, header will be inserted at the last.
        :raise Table.SameHeaderExistError: if there is the same header.
        :raise ValueError: if col_h is None.
        """
        if col_h is None:
            raise ValueError
        try:
            self._header_h.index(col_h)
            raise Table.SameHeaderExistError
        except ValueError:
            pass
        if col is None:
            col = self._n_col
        self.insert_col(col)
        self._header_h.insert(col, col_h)

    def append_header_vs(self, row_hs: List[VHT]):
        """ Appends a list of vertical headers.

        Empty rows will be inserted where the headers are inserted.

        :param row_hs: a list of vertical headers
        :raise Table.SameHeaderExistError: if there are the same headers.
        :raise ValueError: if there is None in the list.
        """
        if row_hs.count(None) is not 0:
            raise ValueError
        for a, b in combinations(row_hs, 2):
            if a == b:
                raise Table.SameHeaderExistError
        for h in row_hs:
            self.insert_header_v(h)

    def append_header_hs(self, col_hs: List[HHT]):
        """ Appends a list of horizontal headers.

        Empty columns will be inserted where the headers are inserted.

        :param col_hs: a list of col_hs headers
        :raise Table.SameHeaderExistError: if there are the same headers.
        :raise ValueError: if there is None in the list.
        """
        if col_hs.count(None) is not 0:
            raise ValueError
        for a, b in combinations(col_hs, 2):
            if a == b:
                raise Table.SameHeaderExistError
        for h in col_hs:
            self.insert_header_h(h)

    def __eq__(self, other):
        if isinstance(other, Table):
            return other.header_v == self._header_v \
                   and other.header_h == self._header_h \
                   and super(Table, self).__eq__(other)
        return super(Table, self).__eq__(other)

    def __str__(self):
        """ Prints a table. """
        col_space = [0] * (self._n_col + 1)
        for v_h in self._header_v:
            if col_space[0] < len(str(v_h)):
                col_space[0] = len(str(v_h))

        for i, h_h in enumerate(self._header_h):
            col_space[i + 1] = len(str(h_h))

        for r in self._matrix:
            for i, cell in enumerate(r):
                if col_space[i + 1] < len(str(cell)):
                    col_space[i + 1] = len(str(cell))

        result = " " * col_space[0] + "    "
        for i, h_h in enumerate(self._header_h):
            result += ("%" + str(col_space[i + 1]) + "s    ") % h_h
        result += "\n"

        for i, r in enumerate(self._matrix):
            result += ("%" + str(col_space[0]) + "s    ") % self._header_v[i]
            for j, cell in enumerate(r):
                result += ("%" + str(col_space[j + 1]) + "s    ") % cell
            result += "\n"
        return result

    class SameHeaderExistError(BaseException):
        """ Raised when the same header is inserted. """

        def __init__(self): pass

        def __str__(self): return "Header must be unique."


if __name__ == "__main__":
    m: Matrix[int] = Matrix[int]()

    m.insert_row()
    m.insert_row()
    m.insert_col()
    m.insert_col()
    print("test insert_row(), insert_col()\n" + str(m))

    m.insert(1, 0, 0)
    m.insert(2, 1, 0)
    m.insert(3, 4, 2)
    print("test insert\n" + str(m))
    print("row: " + str(m.n_row))
    print("col: " + str(m.n_col))
    print("get(4,2): " + str(m.get(4, 2)))
    print("get_row(0): " + str(m.get_row(0)))
    print("get_col(2): " + str(m.get_col(2)))
    print("m[4][2]: " + str(m[4][2]))
    print()

    m[4][2] = 20
    print("m[4][2] after change: " + str(m[4][2]))
    print()

    m.insert_row(1)
    m.insert_col(2)
    print("test insert_row(1), insert_col(2)\n" + str(m))

    m.remove_row(4)
    m.remove_col(0)
    print("test remove_row(4), remove_col(0)\n" + str(m))

    t: Table[int, int, str] = Table[int, int, str]()
    t.append_header_vs([0, 1, 2, 3, 4, 6])
    t.append_header_hs(['aaaaa', 'bbbbbbb', 'ccccccc', 'dddddddd'])
    print("test append_header_vs, append_header_hs\n" + str(t))

    t.insert_header_h('k', 3)
    t.insert_header_v(20, 2)
    print("test insert_header_v, insert_header_h\n" + str(t))

    t.insert_with_header(1000000, 20, 'k')
    print("test insert_with_header\n" + str(t))

    t.insert_header_v(4)  # Table.SameHeaderExistError
    print(t)
