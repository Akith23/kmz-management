# kmz-management

## Project Description

This application allows users to load Keyhole Markup language Zipped (KMZ) files, display the placemarks (pins) on a map, and create routes from selected pins. Users can customize the name and color of the routes and save them as Keyhole Markup Language (KML) files. The application also supports automatic route creation based on the source KMZ file.

## Installation Instructions

- Python 3 is required.
- Install the necessary pip dependencies using the following command:
  ```bash
  pip install lxml simplekml tkintermapview
  ```

## How to Run the Application

Execute the following command in your terminal:

```bash
python AIKC/"Rutas a Puntos"/ruta_por_punto.py
```

## Main Features

- Load KMZ files.
- Display placemarks (pins) on a map.
- Select pins on the map.
- Create routes from selected pins, with custom names and colors.
- Automatically create routes based on the source KMZ file.
- Save generated routes to a KML file.
- Clear the map and loaded data.