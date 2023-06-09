import sys
import functools


def _validate_object(object, const: int, name: str = None) -> None:
    """Check that type of `object` is equal `const`."""
    try:
        value = getattr(object, "Type")
    except AttributeError as e:
        sys.exit("The object doesn't have an attribute 'Type'.")
    else:
        if value == const:
            return
        else:
            if name is None:
                sys.exit(f"The object is not a valid object.")
            else:
                sys.exit(f"The object is not an Inventor API {name} Object.")


class Document:
    _DOCUMENT_TYPES = {"part": 12290, "assembly": 12291, "drawing": 12292}

    _SUBTYPES = {
        "sheet metal": "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}",
        "assembly": "{E60F81E1-49B3-11D0-93C3-7E0706000000}",
        "modeling": "{4D29B490-49B2-11D0-93C3-7E0706000000}",
        "drawing layout": "{BBF9FDF1-52DC-11D0-8C04-0800090BE8EC}",
    }

    _PROPERTY_SETS = (
        "Design Tracking Properties",  # 53 items, main set
        "Inventor Summary Information",  # 8 items
        "Inventor Document Summary Information",  # 3 items
        "Inventor User Defined Properties",  # 0 items, by default
    )

    def __init__(self, document_inventor_api_object):
        _validate_object(document_inventor_api_object, 50332160, "Document")
        self._document = document_inventor_api_object

    @property
    def file(self):
        return self._document

    @property
    def path(self) -> str:
        return self._document.FullFileName

    @property
    def name(self) -> str:
        return self._document.DisplayName

    @property
    def component_definition(self):
        return self._document.ComponentDefinition

    @property
    def unit_quantity(self) -> str:
        try:
            self.component_definition.BOMQuantity.BaseUnits = "м"
        except Exception:
            return ""
        else:
            return self.component_definition.BOMQuantity.UnitQuantity

    def get_properties(self, properties: tuple[str]) -> dict:
        data = dict()

        value = None
        for prop in properties:
            for set_title in Document._PROPERTY_SETS:
                try:
                    value = self._document.PropertySets.Item(set_title).Item(prop).Value
                except Exception:
                    continue
                else:
                    break

            data[prop] = value
        return data

    def update(self):
        self._document.Update2()

    def is_part(self) -> bool:
        return any((self.is_modeling(), self.is_sheet_metal()))

    def is_modeling(self) -> bool:
        return self._document.SubType == Document._SUBTYPES["modeling"]

    def is_sheet_metal(self) -> bool:
        return self._document.SubType == Document._SUBTYPES["sheet metal"]

    def is_assembly(self) -> bool:
        return self._document.SubType == Document._SUBTYPES["assembly"]

    def is_drawing(self) -> bool:
        return self._document.SubType == Document._SUBTYPES["drawing layout"]


class Assembly(Document):
    # Types of BOM views in Inventor API
    BOM_VIEWS = {
        "raw": 62465,
        "structured": 62466,
        "parts_only": 62467,
    }

    def __init__(self, document_inventor_api_object) -> None:
        super().__init__(document_inventor_api_object)

        if not super().is_assembly():
            raise Exception(f"File {self.name} is not an Inventor Assembly.")
        self._bom = document_inventor_api_object.ComponentDefinition.BOM

    @property
    def bom(self):
        return self._bom

    @property
    def raw_bom(self):
        return BOMView(self._bom.BOMViews.Item(1))

    @property
    def structured_bom(self):
        self._bom.StructuredViewEnabled = True
        return BOMView(self._bom.BOMViews.Item(2))

    @property
    def parts_only_bom(self):
        self._bom.PartsOnlyViewEnabled = True
        return BOMView(self._bom.BOMViews.Item(3))


class Part(Document):
    def __init__(self, document_inventor_api_object) -> None:
        super().__init__(document_inventor_api_object)

        if not super().is_part():
            raise Exception(f"File {self.name} in not an Inventor Part.")

    def get_size(self) -> int:
        """
        Return size in millimeters of the part along `axis`: X, Y, Z.
        """
        size = dict()
        for axis in ("X", "Y", "Z"):
            max_p = getattr(self.component_definition.RangeBox.MaxPoint, axis)
            min_p = getattr(self.component_definition.RangeBox.MinPoint, axis)
            size[f"Size {axis}"] = round((max_p - min_p) * 10)

        return size


class Drawing(Document):
    def __init__(self, document_inventor_api_object) -> None:
        super().__init__(document_inventor_api_object)

        if not super().is_drawing():
            raise Exception(f"File {self.name} in not an Inventor DrawingView.")


class BOMView:
    def __init__(self, bom_view_inventor_api_object):
        _validate_object(bom_view_inventor_api_object, 100674304, "BOMView")
        self._bom_view = bom_view_inventor_api_object
        self._rows_count = 1

    def get_rows(self, bom_structures=("normal", "purchased", "phantom")):
        """
        Get list of BOMRow objects with given `bom_structures`.
        """
        rows = []

        for bom_row in self._bom_view.BOMRows:
            row = BOMRow(bom_row)
            valid = True
            if bom_structures and row.bom_structure not in bom_structures:
                valid = False
            if valid:
                rows.append(row)

        return rows

    @functools.cached_property
    def rows_count(self):
        self._number_of_rows(self._bom_view.BOMRows)
        return self._rows_count

    def _number_of_rows(self, rows, count_child_rows=True):
        # Base case
        if rows is None:
            return
        self._rows_count += rows.Count - 1

        if count_child_rows:
            for row in rows:
                self._number_of_rows(row.ChildRows)


class BOMRow:
    _BOM_STRUCTURES = {
        51969: "default",
        51970: "normal",
        51971: "phantom",
        51972: "reference",
        51973: "purchased",
        51974: "inseparable",
        51975: "varies",
    }

    def __init__(self, bom_row_inventor_api_object):
        self._row = bom_row_inventor_api_object

    @property
    def number(self) -> str:
        return self._row.ItemNumber

    @property
    def quantity(self) -> str:
        q = self._row.TotalQuantity
        if q == "Null":
            print(f"Ivalid total quantity in row #{self.number}.")
        return q

    @property
    def item_quantity(self) -> int:
        return self._row.ItemQuantity

    @property
    def bom_structure(self) -> str:
        return BOMRow._BOM_STRUCTURES.get(self._row.BOMStructure)

    @functools.cached_property
    def item(self):
        doc = self._row.ComponentDefinitions.Item(1).Document
        if Document(doc).is_part():
            return Part(doc)
        elif Document(doc).is_assembly():
            return Assembly(doc)
        else:
            return Document(doc)

    def is_purchased(self):
        return self.bom_structure == "purchased"

    def is_phantom(self):
        return self.bom_structure == "phantom"

    def is_normal(self):
        return self.bom_structure == "normal"

    def is_merged(self):
        return self._row.Merged

    def get_child_rows(self):
        rows = []
        if child_rows := self._row.ChildRows:
            for row in child_rows:
                rows.append(BOMRow(row))
        return rows
