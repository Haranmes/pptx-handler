"""Main module."""
from pptx import Presentation as pp
import pandas as pd
from datetime import datetime
from pathlib import Path
import re
import json
import xlwings as xw

class PowerpointHandler:
    """
    A class to handle PowerPoint presentations, including adding tables, charts, and images.

    Attributes:
        costumer_name (str): The name of the customer.
        powerpoint_dir (Path): The directory where the PowerPoint files are stored.
        powerpoint_imges (Path): The directory where the PowerPoint images are stored.
        logo_path (Path): The path to the logo image.
        pp (Presentation): The PowerPoint presentation object.
        elements (dict): A dictionary to store the elements of each slide.
        elements_file_path (Path): The path to the JSON file storing the elements.
        chart_path_with_names (list): A list of paths to the exported charts.
    """

    def __init__(self, powerpoint_dir: Path, powerpoint_images_dir: Path, costumer_name: str, target_dir: Path):
        """
        Initializes the PowerpointHandler with the given directories and customer name.
git
        Args:
            powerpoint_dir (Path): The directory where the PowerPoint files are stored.
            powerpoint_imges (Path): The directory where the PowerPoint images are stored.
            costumer_name (str): The name of the customer.
        """
        template_file_name = "202x-xx-xx_Datenanalyse_AKL_Kundenname.pptx"
        self.powerpoint_dir = powerpoint_dir
        self.powerpoint_imges = powerpoint_images_dir

        # get working directory
        self.template_dir = Path(__file__).resolve().parent / 'template' / template_file_name
        self.elements_file_path = self.powerpoint_dir / 'elements.json'

        self.pp = pp(self.template_dir)
        self.elements = self.__get_elements_per_slide()

        self.costumer_name = costumer_name
        self.target_dir = target_dir
        self.logo_path = None

        for image in self.powerpoint_imges:
            file_path = Path(image)
            if file_path.suffix == '.png':
                file_name = file_path.stem
                if self.like_operator('%ogo%', file_name):
                    self.logo_path = file_path




    def like_operator(self, pattern, string) -> bool:
        """
        Converts SQL LIKE pattern to regex pattern and matches it with the given string.

        Args:
            pattern (str): The SQL LIKE pattern.
            string (str): The string to match.

        Returns:
            bool: True if the string matches the pattern, False otherwise.
        """
        regex_pattern = pattern.replace('%', '.*').replace('_', '.')
        return re.match(regex_pattern, string) is not None

    def __get_elements_per_slide(self) -> dict:
        """
        Extracts the indices of all shapes in the given slide and saves them to a JSON file.

        Returns:
            dict: A dictionary with slide indices as keys and shape indices as values.
        """
        slide_shapes_name = {}
        for slide_idx, slide in enumerate(self.pp.slides):
            slide_shapes_name[slide_idx] = {}
            for shape_idx, shape in enumerate(slide.shapes):
                slide_shapes_name[slide_idx][shape.name] = shape_idx

        with open(self.elements_file_path, 'w') as json_file:
            json.dump(slide_shapes_name, json_file, indent=4)

        print(slide_shapes_name)
        return slide_shapes_name

    def __update_elements_of_slide(self, slide_number: int) -> None:
        """
        Updates the elements of a specific slide and saves the changes to the JSON file.

        Args:
            slide_number (int): The slide number to update.
        """
        slide = self.pp.slides[slide_number] if self.pp.slides[slide_number] is not None else None
        if slide is None:
            raise ValueError("The slide number is not valid.")

        # After

        # for shape_idx, shape in enumerate(slide.shapes):
        #     if self.elements[slide_number].get(shape.name, None) is None:
        #         self.elements[slide_number][shape.name] = shape_idx
        #     elif self.elements[slide_number].get(shape.name, None) != shape_idx:
        #         self.elements[slide_number][shape.name] = shape_idx

        # Add new shapes to self.elements
        for shape_idx, shape in enumerate(slide.shapes):
            if shape.name not in self.elements[slide_number]:
                self.elements[slide_number][shape.name] = shape_idx

        # Remove shapes from self.elements that no longer exist in slide.shapes
        existing_shape_names = {shape.name for shape in slide.shapes}
        for shape_name in list(self.elements[slide_number].keys()):
            if shape_name not in existing_shape_names:
                del self.elements[slide_number][shape_name]


        with open(self.elements_file_path, 'w') as json_file:
            json.dump(self.elements, json_file, indent=4)

    def __get_shape_and_slide(self, slide_number: int, shape_name: str) -> tuple:
        """
        Retrieves the shape and slide for a given slide number and shape name.

        Args:
            slide_number (int): The slide number.
            shape_name (str): The name of the shape.

        Returns:
            tuple: A tuple containing the shape and slide objects.
        """
        slide = self.pp.slides[slide_number] if self.pp.slides[slide_number] is not None else None
        if slide is None:
            raise ValueError("The slide number is not valid.")
        shape_id = self.elements[slide_number].get(shape_name, "Shape not found")
        if shape_name == "Shape not found":
            return None
        shape = slide.shapes[shape_id]
        return shape, slide

    def add_logo(self, image_path: str, shape_name: str = "logo", slide_number: int = 0) -> None:
        """
        Adds a logo image to the slide by replacing an existing shape.

        Args:
            image_path (str): The path to the image file to be added.
            shape_name (str): The name of the shape to be replaced.
            slide_number (int): The slide number where the image will be added.
        """
        shape, slide = self.__get_shape_and_slide(slide_number, shape_name)
        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height
        image_path_string = str(image_path)
        slide.shapes.add_picture(image_path_string, left, top, width, height)

    def add_costumer_name(self, costumer_name: str, slide_number: int = 0, shape_name: str = "costumer") -> None:
        """
        Adds the customer name to the slide by replacing an existing shape.

        Args:
            costumer_name (str): The name of the customer.
            slide_number (int): The slide number where the name will be added.
            shape_name (str): The name of the shape to be replaced.
        """
        shape, slide = self.__get_shape_and_slide(slide_number, shape_name)
        if shape.has_text_frame is not True:
            raise ValueError("The shape is not a text frame.")
        shape.text = costumer_name
        self.__update_elements_of_slide(slide_number)
        self.save_presentation()

    def save_presentation(self):
        """
        Saves the PowerPoint presentation with the current date and customer name.
        """
        current_date = datetime.now().strftime('%Y-%m-%d')
        output_path = f"{current_date}_Datenanalyse_AKL_{self.costumer_name}.pptx"
        self.pp.save(self.target_dir / output_path)
        print(f"Presentation saved to {self.target_dir / output_path}")

    def add_table(self, title: str, slide_number: int, table: pd.DataFrame, shape_name: str):
        """
        Adds a table to the slide.

        Args:
            title (str): The title of the table.
            slide_number (int): The slide number where the table will be added.
            table (pd.DataFrame): The data to be displayed in the table.
            shape_name (str): The name of the shape to be replaced.
        """
        shape, slide = self.__get_shape_and_slide(slide_number, shape_name)

        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height
        rows, cols = table.shape

        sp = shape._element
        sp.getparent().remove(sp)

        table_shape = slide.shapes.add_table(rows + 1, cols, left, top, width, height).table
        for col_idx, col_name in enumerate(table.columns):
            cell = table_shape.cell(0, col_idx)
            cell.text = col_name
        for row_idx, row in table.iterrows():
            for col_idx, value in enumerate(row):
                cell = table_shape.cell(row_idx + 1, col_idx)
                cell.text = str(value)

        self.__update_elements_of_slide(slide_number)
        self.save_presentation()

    def export_plot_from_excel(self, path_to_excel_file: str, sheet_number: int = 0) -> list:
        """
        Exports charts from an Excel file to PNG images.

        Args:
            path_to_excel_file (str): The path to the Excel file.
            sheet_number (int): The sheet number to export charts from.

        Returns:
            list: A list of paths to the exported chart images.
        """
        chart_paths = []
        with xw.App() as app:
            book = app.books.open(path_to_excel_file)
            sheet = book.sheets[sheet_number]
            for chart in sheet.charts:
                chart_name = chart.name
                chart_path = str(self.powerpoint_dir / f"{chart_name}.png")
                chart.to_png(chart_path)
                chart_paths.append(chart_path)
        return chart_paths

    def add_chart_from_file(self, slide_number: int, shape_name: str, chart_path_with_name_and_type: str) -> None:
        """
        Adds a chart image to the slide from a file.

        Args:
            slide_number (int): The slide number where the chart will be added.
            shape_name (str): The name of the shape to be replaced.
            chart_path_with_name_and_type (str): The path to the chart image file.
        """
        shape, slide = self.__get_shape_and_slide(slide_number, shape_name)
        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height

        sp = shape._element
        sp.getparent().remove(sp)

        chart_path = str(self.powerpoint_dir / f"{chart_path_with_name_and_type}")
        slide.shapes.add_picture(chart_path, left, top, width, height)

        self.__update_elements_of_slide(slide_number)
        self.save_presentation()

    def add_table_from_excel(self, slide_number: int, shape_name: str, path_to_excel_file: str, sheet_number: int = 0) -> None:
        """
        Adds a table to the slide from an Excel file.

        Args:
            slide_number (int): The slide number where the table will be added.
            shape_name (str): The name of the shape to be replaced.
            path_to_excel_file (str): The path to the Excel file.
            sheet_number (int): The sheet number to get the table from.
        """
        shape, slide = self.__get_shape_and_slide(slide_number, shape_name)
        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height
        with xw.App() as app:
            book = app.books.open(path_to_excel_file)
            sheet = book.sheets[sheet_number]
            table = sheet.used_range.value
            rows, cols = len(table), len(table[0])
            table_shape = slide.shapes.add_table(rows, cols, left, top, width, height).table
            for row_idx, row in enumerate(table):
                for col_idx, value in enumerate(row):
                    cell = table_shape.cell(row_idx, col_idx)
                    cell.text = str(value)
            book.close()
        self.__update_elements_of_slide(slide_number)
        self.save_presentation()
