import tkinter
from tkinter import ttk, filedialog, messagebox
import zipfile
from io import BytesIO
import os

try:
    from lxml import etree # Para parsear XML (KML)
except ImportError:
    messagebox.showerror("Error de Importación", "La biblioteca lxml no está instalada. Por favor, instálala con 'pip install lxml'")
    exit()

try:
    import simplekml # Para generar el KML de salida
except ImportError:
    messagebox.showerror("Error de Importación", "La biblioteca simplekml no está instalada. Por favor, instálala con 'pip install simplekml'")
    exit()

try:
    import tkintermapview
except ImportError:
    messagebox.showerror("Error de Importación", "La biblioteca tkintermapview no está instalada. Por favor, instálala con 'pip install tkintermapview'")
    exit()

# Namespaces comunes en KML used for parsing KML files.
KML_NS = "{http://www.opengis.net/kml/2.2}"  # KML namespace
GX_NS = "{http://www.google.com/kml/ext/2.2}"  # Google Extensions namespace
ATOM_NS = "{http://www.w3.org/2005/Atom}"  # Atom namespace
NS_MAP = {  # Namespace map for lxml
    'kml': 'http://www.opengis.net/kml/2.2',
    'gx': 'http://www.google.com/kml/ext/2.2',
    'atom': 'http://www.w3.org/2005/Atom'
}

# Color Constants: Defines colors for UI elements and KML output.
# User-facing color names for UI elements (e.g., combobox).
COLOR_RED_NAME = "rojo" 
COLOR_GREEN_NAME = "verde"
COLOR_BLUE_NAME = "azul"
COLOR_CYAN_NAME = "cyan"

# Internal color names used by tkintermapview.
COLOR_RED = "red"
COLOR_GREEN = "green"
COLOR_BLUE = "blue"
COLOR_CYAN = "cyan"

# Default color settings.
DEFAULT_ROUTE_COLOR_UI_NAME = COLOR_RED_NAME # Default color displayed in the UI for routes.
DEFAULT_ROUTE_COLOR_INTERNAL = COLOR_RED # Default internal color for route lines on the map.
DEFAULT_MARKER_COLOR = COLOR_RED # Default color for map markers.
SELECTED_MARKER_COLOR = COLOR_GREEN # Color for selected map markers.

# KML color codes (ABGR format).
KML_COLOR_RED = "ff0000ff"
KML_COLOR_GREEN = "ff00ff00"
KML_COLOR_BLUE = "ffff0000"
KML_COLOR_CYAN = "ffffff00"
DEFAULT_KML_COLOR = KML_COLOR_RED # Default color for KML linestrings.

# Theme Color Dictionaries
DARK_THEME_COLORS = {
    "bg": "#2E2E2E",
    "fg": "#E0E0E0",
    "button_bg": "#505050",
    "button_fg": "#E0E0E0",
    "entry_bg": "#3C3C3C",
    "entry_fg": "#E0E0E0",
    "list_bg": "#3C3C3C",
    "list_fg": "#E0E0E0",
    "label_bg": "#2E2E2E",
    "separator_bg": "#505050",
    "map_bg": "#1C1C1C",
    "frame_bg": "#2E2E2E",
    "paned_bg": "#2E2E2E",
    "checkbutton_bg": "#2E2E2E",
    "checkbutton_fg": "#E0E0E0",
    "checkbutton_select": "#4A4A4A", # Color when checkbutton is selected
    "labelframe_fg": "#E0E0E0", # Foreground for LabelFrame text
}
LIGHT_THEME_COLORS = {
    "bg": "#F0F0F0",
    "fg": "#000000",
    "button_bg": "#E1E1E1",
    "button_fg": "#000000",
    "entry_bg": "#FFFFFF",
    "entry_fg": "#000000",
    "list_bg": "#FFFFFF",
    "list_fg": "#000000",
    "label_bg": "#F0F0F0",
    "separator_bg": "#D3D3D3",
    "map_bg": "#FFFFFF",
    "frame_bg": "#F0F0F0",
    "paned_bg": "#F0F0F0",
    "checkbutton_bg": "#F0F0F0",
    "checkbutton_fg": "#000000",
    "checkbutton_select": "#C0C0C0", # Color when checkbutton is selected
    "labelframe_fg": "#000000", # Foreground for LabelFrame text
}


class KMZRouteApp(tkinter.Tk):
    """
    A tkinter application for loading KMZ files, visualizing placemarks (pins) on a map,
    and creating routes from these pins. Routes can be customized (name, color) and
    saved as KML files. The application also supports automatic route generation based
    on the source KMZ file of the pins.
    """
    def __init__(self):
        """
        Initializes the KMZRouteApp application.

        Sets up the main window title and geometry. Initializes internal data structures
        for storing information about pins (placemarks) and routes. Configures the
        default position and zoom level for the map widget. Calls the `_setup_ui`
        method to create and arrange all user interface elements.
        """
        super().__init__()
        self.title("Visor KMZ con LXML y SimpleKML")
        self.geometry("1200x800")

        self.pins_data = []  # Stores data for each placemark (name, coordinates, selection state, etc.)
        self.routes_data = []  # Stores data for each created route (name, coordinates, color)
        self.map_markers = []  # References to marker objects on the tkintermapview widget
        self.map_paths = []  # References to path objects (routes) on the tkintermapview widget
        self.last_selected_index = None  # Index of the last clicked pin in the list, for shift-selection
        self.order_counter = 1  # Counter to assign order to selected pins
        self.update_ordering_id = None  # ID for tkinter's `after` mechanism, to schedule UI updates
        self.extraction_error_count = 0 # Counter for errors encountered during placemark coordinate extraction
        
        self.theme = "light"  # Initialize theme to light mode
        self.style = ttk.Style() # Initialize ttk.Style for theming ttk widgets

        self._setup_ui() # Initialize all UI components
        self._apply_theme() # Apply the initial theme
        
        # Set initial map position (Asunción, Paraguay) and zoom level.
        self.map_widget.set_position(-25.2637, -57.5759) 
        self.map_widget.set_zoom(5)

    def toggle_theme(self):
        """Switches the application theme between 'light' and 'dark' and applies it."""
        if self.theme == "light":
            self.theme = "dark"
        else:
            self.theme = "light"
        self._apply_theme()

    def _apply_theme(self):
        """Applies the current theme to all relevant UI widgets."""
        # Determine current colors based on theme
        colors = DARK_THEME_COLORS if self.theme == "dark" else LIGHT_THEME_COLORS
        current_theme_name = "dark" if self.theme == "dark" else "light"

        # Apply to root window
        self.configure(bg=colors["bg"])

        # --- Configure ttk styles ---
        # Note: Some styles like TCombobox's dropdown list arrow color are notoriously hard to change
        # and might depend on the underlying Tk theme/version.
        
        self.style.theme_use('default') # Reset to default to remove OS-specific styling that might interfere

        self.style.configure(".", background=colors["bg"], foreground=colors["fg"]) # Global settings
        self.style.configure("TFrame", background=colors["frame_bg"])
        self.style.configure("TLabel", background=colors["label_bg"], foreground=colors["fg"])
        self.style.configure("TButton", background=colors["button_bg"], foreground=colors["button_fg"], padding=3)
        self.style.map("TButton",
                       background=[('active', colors["button_select"]), ('pressed', colors["button_select"])],
                       foreground=[('active', colors["button_fg"])])
        self.style.configure("TEntry", fieldbackground=colors["entry_bg"], foreground=colors["entry_fg"], insertcolor=colors["fg"])
        
        # Combobox styling (listbox part is harder to style directly via ttk.Style for all themes)
        self.style.configure("TCombobox",
                             fieldbackground=colors["entry_bg"],
                             foreground=colors["entry_fg"],
                             selectbackground=colors["entry_bg"], # Background of the selected item in the entry part
                             selectforeground=colors["entry_fg"], # Foreground of the selected item in the entry part
                             background=colors["button_bg"], # Background of the dropdown arrow area
                             arrowcolor=colors["fg"])
        # Attempt to style the dropdown list (may not work on all platforms/Tk versions)
        self.tk.call("option", "add", "*TCombobox*Listbox.background", colors["list_bg"])
        self.tk.call("option", "add", "*TCombobox*Listbox.foreground", colors["list_fg"])
        self.tk.call("option", "add", "*TCombobox*Listbox.selectBackground", colors["button_select"])
        self.tk.call("option", "add", "*TCombobox*Listbox.selectForeground", colors["list_fg"])

        self.style.configure("TScrollbar", background=colors["button_bg"], troughcolor=colors["bg"], arrowcolor=colors["fg"])
        self.style.map("TScrollbar",
                       background=[('active', colors["button_select"])])
        self.style.configure("TPanedwindow", background=colors["paned_bg"])
        self.style.configure("TSeparator", background=colors["separator_bg"])
        
        self.style.configure("TLabelframe",
                             background=colors["frame_bg"],
                             labelmargins=[5, 0, 5, 0], # Add some margin around the label
                             relief="groove", # Keep relief for visibility
                             bordercolor=colors["separator_bg"]) 
        self.style.configure("TLabelframe.Label",
                             background=colors["frame_bg"], # Match LabelFrame background
                             foreground=colors["labelframe_fg"])
        
        # Checkbutton style
        # The selectcolor is the background of the checkbutton square itself when checked.
        # indicatorcolor is the color of the check mark/indicator.
        self.style.configure("TCheckbutton",
                             background=colors["checkbutton_bg"],
                             foreground=colors["checkbutton_fg"],
                             selectcolor=colors["checkbutton_select"] if self.theme == "light" else colors["bg"], # Make square bg match theme
                             indicatormargin=3,
                             padding=2)
        self.style.map("TCheckbutton",
                       background=[('active', colors["checkbutton_bg"])],
                       foreground=[('active', colors["checkbutton_fg"])],
                       indicatorcolor=[("selected", colors["fg"]), # Checkmark color when selected
                                       ("!selected", colors["fg"])], # Checkmark color when not selected (box outline)
                       selectcolor=[('!disabled', colors["checkbutton_select"] if self.theme == "light" else colors["bg"] )] # bg of square
                      )


        # --- Apply to specific non-ttk and container widgets ---
        # These are direct configurations as they are not ttk widgets or need specific handling.
        
        # Pins Canvas (standard tkinter.Canvas)
        self.pins_canvas.configure(bg=colors["list_bg"])
        
        # self.pins_list_frame is a ttk.Frame, ensure it re-styles if it was already created.
        # This is generally handled by the TFrame style, but explicit re-application can be added if needed.
        # self.pins_list_frame.configure(style="TFrame") # Already done in _setup_ui effectively by being a ttk.Frame
        
        # Update dynamically created checkbuttons for pins (they should pick up TCheckbutton style)
        # If they don't, we might need to iterate and reconfigure, but ttk.Style is preferred.
        # for pin_data in self.pins_data:
        #     if "checkbox_widget" in pin_data and pin_data["checkbox_widget"].winfo_exists():
        #         pin_data["checkbox_widget"].configure(style="TCheckbutton") # Re-apply if needed

        # Map Widget Frame (self.map_frame is a ttk.Frame, styled by "TFrame")
        # The TkinterMapView widget itself is external and may not fully respect these themes.
        # Its background is usually map tiles.

        # Force update of all widgets to reflect new styles
        self.update_idletasks()

    def _setup_ui(self):
        """
        Sets up the user interface of the application.

        This method creates and arranges all the UI elements, including frames for layout,
        buttons for actions (load KMZ, create route, save KML, etc.), lists for
        displaying pins, input fields for route naming, and the map widget itself.
        It uses `ttk` themed widgets for a modern look and feel.
        """
        # Main application frame
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(expand=True, fill="both")

        # Paned window to divide left panel (controls) and right panel (map)
        paned_window = ttk.PanedWindow(main_frame, orient="horizontal")
        paned_window.pack(expand=True, fill="both")

        # Left panel for controls
        left_panel = ttk.Frame(paned_window, width=350, padding="5")
        left_panel.pack_propagate(False) # Prevent left panel from shrinking to content
        paned_window.add(left_panel, weight=1) # Add to paned window, allow resizing

        # Button to load KMZ file
        load_button = ttk.Button(left_panel, text="Cargar Archivo KMZ", command=self.load_kmz_file)
        load_button.pack(pady=10, padx=5, fill="x")

        ttk.Separator(left_panel, orient="horizontal").pack(fill="x", pady=5)

        # Frame and scrollable canvas for displaying list of available pins
        pins_list_frame_container = ttk.LabelFrame(left_panel, text="Pines Disponibles", padding="5")
        pins_list_frame_container.pack(expand=True, fill="both", pady=5, padx=5)

        self.pins_canvas = tkinter.Canvas(pins_list_frame_container, borderwidth=0)
        self.pins_list_frame = ttk.Frame(self.pins_canvas) # Frame inside canvas to hold pin checkbuttons
        scrollbar = ttk.Scrollbar(pins_list_frame_container, orient="vertical", command=self.pins_canvas.yview)
        self.pins_canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.pins_canvas.pack(side="left", fill="both", expand=True)
        # Add the pins_list_frame to the canvas
        self.pins_canvas_window = self.pins_canvas.create_window((0, 0), window=self.pins_list_frame, anchor="nw")

        # Configure scrollregion when pins_list_frame size changes
        self.pins_list_frame.bind("<Configure>", lambda e: self.pins_canvas.configure(scrollregion=self.pins_canvas.bbox("all")))
        # Adjust width of internal frame when canvas width changes
        self.pins_canvas.bind('<Configure>', self._on_canvas_configure)

        # Frame for route creation controls
        route_controls_frame = ttk.LabelFrame(left_panel, text="Crear Ruta", padding="5")
        route_controls_frame.pack(fill="x", pady=10, padx=5)

        ttk.Label(route_controls_frame, text="Nombre de la Ruta:").pack(anchor="w", padx=5)
        self.route_name_entry = ttk.Entry(route_controls_frame) # Input for route name
        self.route_name_entry.pack(fill="x", padx=5, pady=(0,5))
        
        ttk.Label(route_controls_frame, text="Color de la Ruta:").pack(anchor="w", padx=5)
        self.route_color_combo_values = [COLOR_CYAN_NAME, COLOR_RED_NAME, COLOR_GREEN_NAME, COLOR_BLUE_NAME]
        self.route_color_combo = ttk.Combobox(route_controls_frame, values=self.route_color_combo_values, state="readonly") # Combobox for route color
        try:
            # Set default color in combobox
            default_color_index = self.route_color_combo_values.index(DEFAULT_ROUTE_COLOR_UI_NAME)
            self.route_color_combo.current(default_color_index)
        except ValueError:
            self.route_color_combo.current(0) # Default to first color if not found
        self.route_color_combo.pack(fill="x", padx=5, pady=(0,5))
        # Bind color change event to automatically create route if pins are selected
        self.route_color_combo.bind("<<ComboboxSelected>>", self.on_color_change)
        
        # Button to create route from currently selected pins
        create_route_button = ttk.Button(route_controls_frame, text="Crear Ruta con Pines Seleccionados", command=self.create_route_from_selection)
        create_route_button.pack(pady=5, fill="x", padx=5)
        
        # Button to automatically create routes based on KMZ source
        auto_routes_button = ttk.Button(route_controls_frame, text="Crear Rutas Automáticas", command=self.create_routes_from_all)
        auto_routes_button.pack(pady=5, fill="x", padx=5)

        # Frame for multiple selection buttons ("Select All", "Deselect All")
        select_buttons_frame = ttk.Frame(route_controls_frame)
        select_buttons_frame.pack(fill="x", pady=5, padx=5)
        
        select_all_button = ttk.Button(select_buttons_frame, text="Seleccionar Todos", command=self.select_all_pins)
        select_all_button.pack(side="left", expand=True, fill="x", padx=(0,2))
        deselect_all_button = ttk.Button(select_buttons_frame, text="Deseleccionar Todos", command=self.deselect_all_pins)
        deselect_all_button.pack(side="left", expand=True, fill="x", padx=(2,0))

        # Button to save generated routes to a KML file
        save_routes_button = ttk.Button(left_panel, text="Guardar Rutas Generadas (KML con SimpleKML)", command=self.save_routes_to_kml)
        save_routes_button.pack(pady=10, padx=5, fill="x")
        
        # Button to clear all data (pins, routes) from the map and application
        clear_map_button = ttk.Button(left_panel, text="Limpiar Mapa (Pines y Rutas)", command=self.clear_map_and_data)
        clear_map_button.pack(pady=5, padx=5, fill="x")

        # Theme toggle button
        theme_button = ttk.Button(left_panel, text="Toggle Theme", command=self.toggle_theme)
        theme_button.pack(pady=10, padx=5, fill="x")

        # Right panel for the map widget
        # Storing map_frame as an instance variable for potential direct styling if needed
        self.map_frame = ttk.Frame(paned_window, padding="5") 
        paned_window.add(self.map_frame, weight=3) # Add to paned window, allow resizing

        # TkinterMapView widget
        self.map_widget = tkintermapview.TkinterMapView(self.map_frame, corner_radius=0)
        self.map_widget.pack(expand=True, fill="both")

    def _on_canvas_configure(self, event):
        paned_window.add(map_frame, weight=3) # Add to paned window, allow resizing

        # TkinterMapView widget
        self.map_widget = tkintermapview.TkinterMapView(map_frame, corner_radius=0)
        self.map_widget.pack(expand=True, fill="both")

    def _on_canvas_configure(self, event):
        """
        Handles the configure event for the pins canvas.

        Adjusts the width of the frame inside the canvas to match the canvas width,
        ensuring the scrollbar behaves correctly and items in the list don't get
        cut off horizontally.

        Args:
            event: The event object containing details about the configure event,
                   including the new width of the canvas.
        """
        canvas_width = event.width
        self.pins_canvas.itemconfig(self.pins_canvas_window, width=canvas_width)

    def _clear_pin_list_ui(self):
        """
        Clears all checkbutton widgets from the pin list UI frame.
        This is typically called before loading a new KMZ or clearing all data.
        """
        for widget in self.pins_list_frame.winfo_children():
            widget.destroy()

    def _clear_map_markers(self):
        """
        Removes all markers from the `tkintermapview` widget and clears the
        internal list `self.map_markers` that stores references to them.
        """
        for marker in self.map_markers:
            marker.delete()
        self.map_markers = []

    def _clear_map_paths(self):
        """
        Removes all paths (routes) from the `tkintermapview` widget and clears
        the internal list `self.map_paths` that stores references to them.
        """
        for path in self.map_paths:
            path.delete()
        self.map_paths = []

    def clear_map_and_data(self):
        """
        Clears all loaded data, UI elements related to pins and routes, and map features.

        Resets the application to a near-initial state by:
        - Clearing the list of pin checkbuttons in the UI.
        - Removing all markers from the map.
        - Removing all paths (routes) from the map.
        - Clearing internal data storage for pins (`self.pins_data`) and routes (`self.routes_data`).
        - Resetting the route name entry field to be empty.
        - Resetting the map zoom to its default overview level.
        Finally, it shows an informational message to the user.
        """
        self._clear_pin_list_ui()
        self._clear_map_markers()
        self._clear_map_paths()
        self.pins_data = []
        self.routes_data = []
        self.route_name_entry.delete(0, tkinter.END) # Clear route name input
        self.map_widget.set_zoom(5) # Reset map zoom
        messagebox.showinfo("Limpieza Completa", "Se han eliminado todos los pines y rutas del mapa y la aplicación.")

    def load_kmz_file(self):
        """
        Loads placemark (pin) data from a KMZ file selected by the user.

        Steps:
        1.  Opens a file dialog for the user to select a `.kmz` file.
        2.  If a file is selected, it first calls `clear_map_and_data` to reset the current state.
        3.  Stores the base name of the selected file in `self.current_source`.
        4.  Opens the KMZ (which is a zip archive) and looks for the first `.kml` file within it.
        5.  Reads the content of the KML file.
        6.  Parses the KML content using `lxml.etree`. An `XMLParser` is configured to
            disable entity resolution for security, keep CDATA content, and remove XML comments.
        7.  Resets `self.pins_data` and `self.extraction_error_count`.
        8.  Calls `self._extract_placemarks_from_lxml_tree` to recursively find and process
            Placemark elements from the parsed KML.
        9.  Displays messages to the user indicating how many pins were loaded and if any
            were skipped due to coordinate format errors.
        10. If pins were loaded, it calls `self._populate_pin_list_ui` to update the UI
            and `self._zoom_to_pins` to adjust the map view.
        11. Handles potential exceptions during the process (e.g., invalid KMZ/KML,
            file I/O errors) and shows an error message.
        """
        filepath = filedialog.askopenfilename(
            title="Seleccionar Archivo KMZ",
            filetypes=(("Archivos KMZ", "*.kmz"), ("Todos los archivos", "*.*"))
        )
        if not filepath: # User cancelled the dialog
            return

        self.clear_map_and_data() # Clear existing data before loading new file
        # Store the name of the loaded KMZ file to identify the source of pins
        self.current_source = os.path.basename(filepath)

        try:
            # Open the KMZ (zip) file
            with zipfile.ZipFile(filepath, 'r') as kmz:
                kml_filename = None
                # Find the KML file within the KMZ archive (usually doc.kml or a single .kml file)
                for name in kmz.namelist():
                    if name.lower().endswith('.kml'):
                        kml_filename = name
                        break
                
                if not kml_filename:
                    messagebox.showerror("Error en KMZ", "No se encontró un archivo KML dentro del KMZ.")
                    return

                # Read the KML file content from the KMZ archive
                kml_bytes = kmz.read(kml_filename)
            
            # Parse the KML content using lxml
            # Options: resolve_entities=False for security, strip_cdata=False to keep CDATA content, remove_comments=True to ignore KML comments
            parser = etree.XMLParser(resolve_entities=False, strip_cdata=False, remove_comments=True)
            xml_root = etree.fromstring(kml_bytes, parser=parser) # Get the root element of the KML
            
            self.pins_data = [] # Reset internal list of pins
            self.extraction_error_count = 0 # Reset error counter for this file load
            # Recursively extract placemarks from the parsed KML tree
            self._extract_placemarks_from_lxml_tree(xml_root)

            num_loaded = len(self.pins_data)
            num_skipped = self.extraction_error_count
            source_name = self.current_source

            # Display feedback to the user about the loading process
            if num_loaded > 0:
                success_msg = f"Se cargaron {num_loaded} pines desde {source_name}."
                if num_skipped > 0:
                    skipped_msg = f" Se omitieron {num_skipped} pines debido a errores en el formato de coordenadas."
                    messagebox.showinfo("KMZ Cargado Parcialmente", success_msg + skipped_msg)
                else:
                    messagebox.showinfo("KMZ Cargado", success_msg)
                self._populate_pin_list_ui() # Update the UI list of pins
                self._zoom_to_pins() # Adjust map view to show loaded pins
            else: # No pins were successfully loaded
                if num_skipped > 0:
                    messagebox.showwarning("Error de Carga de Pines", f"No se cargaron pines desde {source_name}. Se omitieron {num_skipped} pines debido a errores en el formato de coordenadas.")
                else: # No pins found and no errors, likely an empty KML or no Point placemarks
                    messagebox.showinfo("Información", f"No se encontraron pines (Placemarks con Puntos) en el archivo KMZ '{source_name}'.")

        except Exception as e: # Catch-all for other potential errors (zip issues, lxml parsing errors)
            messagebox.showerror("Error al Cargar KMZ", f"Ocurrió un error: {e}")
            import traceback # For debugging, print stack trace to console
            print(traceback.format_exc()) 

    def _extract_placemarks_from_lxml_tree(self, xml_element):
        """
        Recursively extracts Placemark elements with Point geometry from an lxml tree.

        This method traverses the KML structure (Document, Folder, Placemark).
        When a Placemark containing a Point is found, it extracts its name and
        coordinates. The coordinates are stored in two formats: `coords_original`
        (lon, lat, alt) as found in the KML, and `coords_map` (lat, lon) for use
        with `tkintermapview`. Each pin's data is stored as a dictionary in
        `self.pins_data`. It also records the source KMZ file for each pin.
        If coordinate parsing fails, `self.extraction_error_count` is incremented.

        Args:
            xml_element: The lxml element to start parsing from (e.g., the root
                         of the KML document or a Folder element).
        """
        for child in xml_element:
            # Skip XML comments
            if isinstance(child, etree._Comment):
                continue
            
            # If the element is a Document or Folder, recurse into it
            if child.tag == f"{KML_NS}Document" or child.tag == f"{KML_NS}Folder":
                self._extract_placemarks_from_lxml_tree(child)
            
            # If the element is a Placemark
            elif child.tag == f"{KML_NS}Placemark":
                placemark_name_element = child.find(f"{KML_NS}name")
                # Use "Pin sin nombre" if name tag is missing or empty
                placemark_name = placemark_name_element.text if placemark_name_element is not None and placemark_name_element.text else "Pin sin nombre"
                
                # Find a Point geometry within the Placemark (can be nested)
                point_element = child.find(f".//{KML_NS}Point") # ".//" searches current element and all descendants
                
                # Skip if no Point geometry is found in this Placemark
                if point_element is None:
                    continue

                coordinates_element = point_element.find(f"{KML_NS}coordinates")
                # Skip if no coordinates tag or if it's empty
                if coordinates_element is None or not coordinates_element.text:
                    continue

                coords_str = coordinates_element.text.strip()
                try:
                    # KML coordinates are typically lon,lat,alt
                    lon_str, lat_str, *alt_str = coords_str.split(',')
                    lon = float(lon_str)
                    lat = float(lat_str)
                    alt = float(alt_str[0]) if alt_str else 0.0 # Altitude is optional, default to 0

                    pin_info = {
                        "name": placemark_name,
                        "coords_original": (lon, lat, alt), # (lon, lat, alt) for KML
                        "coords_map": (lat, lon), # (lat, lon) for tkintermapview
                        "tk_var": tkinter.BooleanVar(value=False), # Selection state for UI checkbox
                        # Store the source KMZ filename for grouping/identification
                        "source": getattr(self, "current_source", "Sin Fuente") # Default if source not set
                    }
                    self.pins_data.append(pin_info)
                except ValueError: # Handle cases where coordinate string is malformed
                    # If coordinates are malformed, skip this placemark and count error
                    self.extraction_error_count += 1
                    pass # Continue to the next placemark/child

    def _populate_pin_list_ui(self):
        """
        Populates the scrollable list in the UI with checkbuttons for each loaded pin
        and places corresponding markers on the map.

        This method first clears any existing pins from the UI list and map markers.
        Then, for each pin dictionary in `self.pins_data`:
        1.  A `ttk.Checkbutton` is created in the `self.pins_list_frame`.
        2.  The checkbutton's selection state is tied to the pin's `tk_var`.
        3.  A click event (`<Button-1>`) on the checkbutton is bound to `self.on_checkbutton_click`
            to handle selection logic (including Shift-click range selection).
        4.  A trace is added to the `tk_var` to call `self.schedule_update_ordering`
            whenever the pin's selection state changes, which updates the displayed order number.
        5.  A marker is placed on the `self.map_widget` at the pin's coordinates.
        6.  The marker's click command is set to `self._on_marker_click` to toggle selection.
        7.  References to the checkbutton widget and map marker are stored in the pin's dictionary.
        Finally, it updates the scroll region of the pins canvas.
        """
        self._clear_pin_list_ui() # Remove old checkbuttons
        self._clear_map_markers() # Remove old map markers

        for i, pin in enumerate(self.pins_data):
            # Create a checkbutton for each pin in the scrollable list
            cb = ttk.Checkbutton(self.pins_list_frame, text=pin["name"], variable=pin["tk_var"])
            cb.pack(anchor="w", fill="x", padx=5)
            pin["checkbox_widget"] = cb # Store reference to the widget
            
            # Bind left-click to handle selection, including Shift-click for range selection
            cb.bind("<Button-1>", lambda event, index=i: self.on_checkbutton_click(event, index))
            # When the checkbutton state changes (tk_var changes), schedule an update to the ordering display
            pin["tk_var"].trace_add("write", lambda *args: self.schedule_update_ordering())

            # Add a marker on the map for the pin
            marker = self.map_widget.set_marker(
                pin["coords_map"][0],  # Latitude
                pin["coords_map"][1],  # Longitude
                text=pin["name"],      # Text displayed with marker (can be None)
                command=lambda m, p=pin: self._on_marker_click(p) # Command to execute when marker is clicked
            )
            self.map_markers.append(marker) # Keep track of map markers
            pin["map_marker"] = marker # Store reference to the marker in the pin data
        
        # Update the scrollable area of the canvas after adding all checkbuttons
        self.pins_list_frame.update_idletasks() # Ensure frame size is calculated
        self.pins_canvas.config(scrollregion=self.pins_canvas.bbox("all"))
        
        # Apply theme to newly created checkbuttons
        self._apply_theme()


    def on_checkbutton_click(self, event, index):
        """
        Handles click events on pin checkbuttons in the list for selection.

        This method implements standard click behavior and Shift-click range selection.
        -   If Shift is pressed and a previous pin was selected (`self.last_selected_index`
            is not None), it determines the range of pins between the last selected
            pin and the currently clicked pin.
        -   The selection state (checked/unchecked) of all pins in this range is set
            to the *intended* new state of the currently clicked pin (if it was unchecked,
            all become checked, and vice-versa).
        -   `self.last_selected_index` is updated to the current pin's index.
        -   If Shift selection is performed, it returns "break" to prevent Tkinter's
            default checkbutton behavior, as the state has already been managed.
        -   If Shift is not pressed, it simply updates `self.last_selected_index`.

        Args:
            event: The Tkinter event object, containing information about the event
                   (e.g., whether Shift key was pressed via `event.state`).
            index: The index of the clicked pin in the `self.pins_data` list.

        Returns:
            "break" (str) if Shift-click range selection was handled, to stop
            further event propagation. Otherwise, implicitly returns None.
        """
        # Check if the Shift key is pressed (mask 0x0001 for Shift)
        SHIFT_MASK = 0x0001 
        if event.state & SHIFT_MASK and self.last_selected_index is not None:
            start = min(self.last_selected_index, index)
            end = max(self.last_selected_index, index)
            
            # Determine the new state based on the pin being clicked.
            # If the current pin's tk_var is False (it's about to be checked), new_state is True.
            # If the current pin's tk_var is True (it's about to be unchecked), new_state is False.
            # This logic ensures that clicking on an unchecked box (while holding shift) selects the range,
            # and clicking on a checked box (while holding shift) deselects the range.
            # Note: The tk_var of pins_data[index] hasn't updated yet from this click.
            # So, if it's currently False, it means the click intends to make it True.
            current_pin_tk_var = self.pins_data[index]["tk_var"]
            new_state = not current_pin_tk_var.get() # This will be the state *after* the click if not for "break"

            for i in range(start, end + 1):
                self.pins_data[i]["tk_var"].set(new_state)
            
            self.last_selected_index = index # Update the last selected index
            return "break" # Prevent default checkbutton behavior as we've handled it
        else:
            # Normal click (no Shift), just update the last selected index
            self.last_selected_index = index

    def _on_marker_click(self, pin_info):
        """
        Handles click events on map markers.

        When a map marker is clicked, this method toggles the selection state
        of the associated pin (by flipping the value of `pin_info["tk_var"]`).
        It then calls `self.update_marker_color` to visually update the marker
        on the map to reflect its new selection state (e.g., change color).
        The `tk_var` change will also trigger `schedule_update_ordering` due to the trace.

        Args:
            pin_info: The dictionary from `self.pins_data` corresponding to the
                      clicked marker. This dictionary contains the `tk_var` and
                      other pin details.
        """
        # Toggle the selection state of the pin's BooleanVar
        pin_info["tk_var"].set(not pin_info["tk_var"].get())
        # Update the marker's visual appearance (color) to reflect the new state.
        # The trace on tk_var will handle updating the order display in the list.
        self.update_marker_color(pin_info)

    def _zoom_to_pins(self):
        """
        Adjusts the map's viewport to encompass all currently loaded pins.

        -   If no map markers exist, it does nothing.
        -   If there is exactly one marker, it centers the map on that marker's
            position and sets a fixed zoom level (e.g., 15).
        -   If there are multiple markers, it calculates a bounding box that
            encloses all marker positions and then uses `self.map_widget.fit_bounding_box`
            to adjust the map's zoom and position to show all pins.
        """
        if not self.map_markers:
            return # No markers to zoom to
        
        if len(self.map_markers) == 1: 
            # Single marker: center on it and set a specific zoom level
            marker = self.map_markers[0]
            self.map_widget.set_position(marker.position[0], marker.position[1])
            self.map_widget.set_zoom(15) 
            return

        # Multiple markers: fit map to their bounding box
        marker_positions = [marker.position for marker in self.map_markers if marker.position is not None]
        if marker_positions:
            lats = [pos[0] for pos in marker_positions]
            lons = [pos[1] for pos in marker_positions]
            # Determine the top-left and bottom-right coordinates of the bounding box
            top_left = (max(lats), min(lons))       # Max latitude, Min longitude
            bottom_right = (min(lats), max(lons))   # Min latitude, Max longitude
            self.map_widget.fit_bounding_box(top_left, bottom_right)

    def create_route_from_selection(self):
        """
        Creates a new route from the currently selected (checked) pins.

        -   It gathers all pins whose `tk_var` is True.
        -   If fewer than two pins are selected, it shows a warning and returns.
        -   The route name is taken from `self.route_name_entry`. If empty, a default
            name (e.g., "Ruta-1") is generated and also populated back into the entry field.
        -   The route color is determined by the selection in `self.route_color_combo`.
            User-facing color names (e.g., "rojo") are mapped to internal color
            values (e.g., "red") for `tkintermapview`.
        -   The original KML coordinates (lon, lat, alt) and map coordinates (lat, lon)
            for the selected pins are collected. The order of pins in the route
            is determined by their current selection order (maintained by `update_ordering`).
        -   A new route dictionary containing the name, KML coordinates, and color
            is appended to `self.routes_data`.
        -   A path is drawn on the `self.map_widget` using the map coordinates and color.
            This path object is stored in `self.map_paths`.
        -   A confirmation message is displayed.
        -   The route name entry field is cleared for the next route.
        """
        # Get pins that are currently selected and sorted by their selection order
        selected_pins_ordered = sorted(
            [pin for pin in self.pins_data if pin["tk_var"].get() and pin.get("select_order") is not None],
            key=lambda p: p["select_order"]
        )

        if len(selected_pins_ordered) < 2:
            messagebox.showwarning("Selección Insuficiente", "Seleccione al menos dos pines ordenados para crear una ruta.")
            return

        route_name = self.route_name_entry.get().strip()
        if not route_name: 
            route_name = f"Ruta-{len(self.routes_data) + 1}" # Generate default name
            self.route_name_entry.insert(0, route_name) # Update UI with generated name

        # Get selected color from combobox, default if empty
        route_color_ui_name = self.route_color_combo.get().strip() or DEFAULT_ROUTE_COLOR_UI_NAME
        
        # Map user-facing color names (e.g., "rojo") to internal tkintermapview color names (e.g., "red")
        ui_to_internal_color_mapping = {
            COLOR_CYAN_NAME: COLOR_CYAN,
            COLOR_RED_NAME: COLOR_RED,
            COLOR_GREEN_NAME: COLOR_GREEN,
            COLOR_BLUE_NAME: COLOR_BLUE,
        }
        route_color_mapped = ui_to_internal_color_mapping.get(route_color_ui_name, DEFAULT_ROUTE_COLOR_INTERNAL)
        
        # Collect coordinates for the route based on the ordered selection
        # route_kml_coords are (lon, lat, alt) for saving to KML
        route_kml_coords = [pin["coords_original"] for pin in selected_pins_ordered]
        # map_coords_list are (lat, lon) for displaying on tkintermapview
        map_coords_list = [pin["coords_map"] for pin in selected_pins_ordered]

        # Store route data internally
        self.routes_data.append({
            "name": route_name,
            "kml_coords": route_kml_coords,
            "color": route_color_mapped # Store the internal color name
        })

        # Draw the route on the map
        map_path = self.map_widget.set_path(map_coords_list, color=route_color_mapped, width=3)
        self.map_paths.append(map_path) # Keep track of map paths

        messagebox.showinfo("Ruta Creada", f"Ruta '{route_name}' creada con {len(selected_pins_ordered)} puntos y añadida al mapa.")
        # Clear the route name field so a new route doesn't reuse the old name by default
        self.route_name_entry.delete(0, tkinter.END)
        self._apply_theme() # Re-apply theme in case message box changed focus or styling

    def save_routes_to_kml(self):
        """
        Saves all created routes to a KML file using the `simplekml` library.

        -   If no routes are present in `self.routes_data`, it shows an info message and returns.
        -   Prompts the user to select a file path and name for saving the KML file
            using a standard save file dialog. If the user cancels, it returns.
        -   Creates a `simplekml.Kml` object.
        -   Maps internal color names (e.g., "red") to KML color codes (ABGR format, e.g., "ff0000ff").
        -   For each route in `self.routes_data`:
            -   A new linestring is added to the KML object using the route's name and
              `kml_coords` (which are in lon, lat, alt order).
            -   The style of the linestring is set, including color (using the mapped KML color)
              and width.
        -   Attempts to save the KML object to the specified file path.
        -   Shows a success or error message.
        """
        if not self.routes_data:
            messagebox.showinfo("Sin Rutas", "No hay rutas creadas para guardar.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Guardar Rutas como KML (con SimpleKML)",
            defaultextension=".kml",
            filetypes=(("Archivos KML", "*.kml"), ("Todos los archivos", "*.*"))
        )
        if not filepath: # User cancelled save dialog
            return

        kml_output = simplekml.Kml(name="Rutas Generadas") # Create a KML object
        
        # Map internal color names (used by tkintermapview) to KML color codes (ABGR format)
        internal_to_kml_color_mapping = {
            COLOR_RED: KML_COLOR_RED,
            COLOR_GREEN: KML_COLOR_GREEN,
            COLOR_BLUE: KML_COLOR_BLUE,
            COLOR_CYAN: KML_COLOR_CYAN,
        }

        for route_info in self.routes_data:
            route_name = route_info["name"]
            coords = route_info["kml_coords"] # These are (lon, lat, alt)
            linestring = kml_output.newlinestring(name=route_name, coords=coords)
            
            # Get the internal color name, default if not found
            route_color_internal_name = route_info.get("color", DEFAULT_ROUTE_COLOR_INTERNAL)
            # Get the KML color code, default if internal name not in map
            kml_color_code = internal_to_kml_color_mapping.get(route_color_internal_name, DEFAULT_KML_COLOR)
            
            linestring.style.linestyle.color = kml_color_code
            linestring.style.linestyle.width = 3 # Set line width
        
        try:
            kml_output.save(filepath) # Save the KML file
            messagebox.showinfo("Guardado Exitoso", f"Rutas guardadas en '{os.path.basename(filepath)}' usando SimpleKML.")
        except Exception as e:
            messagebox.showerror("Error al Guardar con SimpleKML", f"No se pudo guardar el archivo KML: {e}")

    def select_all_pins(self):
        """
        Selects all pins currently loaded in `self.pins_data`.

        Iterates through each pin and sets its associated `tk_var` (Tkinter BooleanVar)
        to `True`. This action checks the corresponding checkbutton in the UI list.
        The change in `tk_var` will also trigger the `schedule_update_ordering` method
        due to the trace, which will then update the display order and marker colors.
        """
        for pin in self.pins_data:
            pin["tk_var"].set(True)

    def deselect_all_pins(self):
        """
        Deselects all pins currently loaded in `self.pins_data`.

        Iterates through each pin and sets its associated `tk_var` (Tkinter BooleanVar)
        to `False`. This action unchecks the corresponding checkbutton in the UI list.
        The change in `tk_var` will also trigger the `schedule_update_ordering` method
        due to the trace, which will then update (clear) the display order and
        reset marker colors.
        """
        for pin in self.pins_data:
            pin["tk_var"].set(False)

    def update_ordering(self):
        """
        Updates the displayed order of selected pins and their marker colors.

        This method is typically called via `schedule_update_ordering` when pin
        selection changes. It performs the following:
        1.  Iterates through all `self.pins_data`:
            -   If a pin is selected (`tk_var` is True) and doesn't have a `select_order`,
                it assigns the current `self.order_counter` and increments the counter.
            -   If a pin is not selected, `select_order` is set to `None`.
        2.  Resets `self.order_counter` to 1 if no pins are currently selected, ensuring
            the next selection sequence starts from 1.
        3.  Creates a list of selected pins and sorts them by their `select_order`.
        4.  Updates the text of each selected pin's checkbutton in the UI to show its
            1-based order (e.g., "1. Pin Name").
        5.  For any pin that is not selected or has its order cleared, its checkbutton
            text is reset to its base name (without the order prefix).
        6.  Calls `update_marker_color` for every pin to reflect its current selection
            status on the map.
        """
        selected_pins_for_ordering = [] # List to hold pins that are currently selected
        for pin in self.pins_data:
            if pin["tk_var"].get(): # If the pin's checkbox is checked
                # Assign an order number if it's newly selected in this sequence
                if "select_order" not in pin or pin["select_order"] is None:
                    pin["select_order"] = self.order_counter
                    self.order_counter += 1
                selected_pins_for_ordering.append(pin)
            else:
                # Clear order if pin is deselected
                pin["select_order"] = None 

        # If no pins are selected at all, reset the main order counter for the next selection sequence
        if not any(p["tk_var"].get() for p in self.pins_data):
            self.order_counter = 1
        
        # Sort the selected pins by their assigned order number for correct display
        # Pins without a select_order (should not happen if selected) are put at the end
        selected_pins_for_ordering.sort(key=lambda p: p.get("select_order", float('inf')))

        # Update the text of the checkbuttons in the UI list
        # For selected pins, prepend the order number (1-based from the sorted list)
        for i, pin_ordered in enumerate(selected_pins_for_ordering):
             if pin_ordered["tk_var"].get() and pin_ordered.get("select_order") is not None:
                base_name = pin_ordered["name"]
                order_prefix = f"{i + 1}. " # Display order is 1-based
                # Ensure checkbox widget exists before trying to configure it
                if "checkbox_widget" in pin_ordered and pin_ordered["checkbox_widget"].winfo_exists():
                    pin_ordered["checkbox_widget"].config(text=order_prefix + base_name)

        # Reset text for deselected pins and update marker colors for all pins
        for pin in self.pins_data:
            # If pin is not selected OR somehow its order was cleared but it's still marked as selected (cleanup)
            if not pin["tk_var"].get() or pin.get("select_order") is None: 
                if "checkbox_widget" in pin and pin["checkbox_widget"].winfo_exists():
                     pin["checkbox_widget"].config(text=pin["name"]) # Reset to base name without order prefix
            self.update_marker_color(pin) # Update marker color based on selection state

    def update_marker_color(self, pin):
        """
        Updates the color of a specific pin's map marker based on its selection state.

        -   Determines the `new_color` for the marker: `SELECTED_MARKER_COLOR` if the
            pin's `tk_var` is True (selected), otherwise `DEFAULT_MARKER_COLOR`.
        -   If the pin already has a `map_marker` object associated with it, that
            existing marker is deleted from the map.
        -   A new marker is created on the `self.map_widget` at the pin's coordinates
            (`pin["coords_map"]`), using its name as text and the `new_color`.
        -   The click command for this new marker is re-bound to `self._on_marker_click`
            to ensure it remains interactive.
        -   The reference to this new marker object is stored back in `pin["map_marker"]`,
            replacing the old one.

        Args:
            pin: The pin dictionary from `self.pins_data`. This dictionary must
                 contain `tk_var` (BooleanVar for selection state), `coords_map`
                 (tuple of lat, lon), `name` (str), and `map_marker` (can be None
                 or an existing marker object).
        """
        new_color = SELECTED_MARKER_COLOR if pin["tk_var"].get() else DEFAULT_MARKER_COLOR
        # Delete the old marker if it exists
        if "map_marker" in pin and pin["map_marker"]:
            pin["map_marker"].delete() 
        
        # Create a new marker with the updated color and re-bind its click command
        new_marker = self.map_widget.set_marker(
            pin["coords_map"][0], # Latitude
            pin["coords_map"][1], # Longitude
            text=pin["name"],
            marker_color_circle=new_color, # Set the circle color for the marker
            command=lambda m, p=pin: self._on_marker_click(p) # Re-bind click command
        )
        pin["map_marker"] = new_marker # Store reference to the new marker

    def schedule_update_ordering(self):
        """
        Schedules a call to `self.update_ordering` to occur after a short delay (100ms).

        This method is crucial for performance and UI responsiveness. When multiple
        pins are selected or deselected rapidly (e.g., via Shift-click, "Select All",
        or programmatic changes to `tk_var`), each change to a `tk_var` (due to the
        trace) would normally trigger an immediate call to `update_ordering`.
        This can lead to many redundant UI updates.

        By scheduling the update:
        -   If an update is already scheduled (`self.update_ordering_id` is not None),
            it cancels the previously scheduled update.
        -   It then schedules `self.update_ordering` to run after 100 milliseconds
            using `self.after()`.
        This effectively "debounces" the updates, ensuring that `update_ordering`
        only runs once after a burst of selection changes.
        """
        if self.update_ordering_id is not None:
            self.after_cancel(self.update_ordering_id) # Cancel any pending update
        # Schedule update_ordering to run after 100ms
        self.update_ordering_id = self.after(100, self.update_ordering) 

    def create_routes_from_all(self):
        """
        Automatically creates routes by grouping all loaded pins by their 'source'
        (the name of the KMZ file they were loaded from).

        -   It iterates through `self.pins_data` and groups pins based on the
            `"source"` key in each pin's dictionary.
        -   For each group (source file):
            -   If the group contains fewer than two pins, it's skipped (a route needs at least two points).
            -   A route name is generated based on the source file name (e.g., "Ruta example.kmz").
            -   The pins within the group are used in the order they were originally loaded
                from the KMZ file. If a specific order within the source is needed,
                `pins_in_group` would need to be sorted here before coordinate extraction.
            -   KML coordinates (`coords_original`) and map coordinates (`coords_map`) are
                collected for the pins in the group.
            -   `DEFAULT_ROUTE_COLOR_INTERNAL` is used as the color for these automatic routes.
            -   The new route data (name, KML coordinates, color) is appended to `self.routes_data`.
            -   A path is drawn on the `self.map_widget` for the new route, and its reference
                is stored in `self.map_paths`.
            -   A counter for created routes is incremented.
        -   Finally, a message box displays the total number of automatic routes created.
        """
        groups = {} # Dictionary to store pins grouped by their source KMZ file
        for pin in self.pins_data:
            src = pin.get("source", "Sin Fuente") # Get source, default if missing
            groups.setdefault(src, []).append(pin) # Add pin to its source group
        
        routes_created_count = 0 
        for src, pins_in_group in groups.items(): # Iterate through each source and its pins
            if len(pins_in_group) < 2: # Need at least two pins to form a route
                continue 
            
            route_name = f"Ruta {src}" # Default name for auto-generated route
            
            # Pins are used in the order they were originally extracted from the KML.
            # If a specific order is needed (e.g., by name or original KML order if not preserved),
            # pins_in_group should be sorted here before extracting coordinates.
            route_kml_coords = [p["coords_original"] for p in pins_in_group]
            map_coords_list = [p["coords_map"] for p in pins_in_group]
            
            route_color_for_auto_route = DEFAULT_ROUTE_COLOR_INTERNAL # Use default color
            
            # Store route data
            self.routes_data.append({
                "name": route_name,
                "kml_coords": route_kml_coords,
                "color": route_color_for_auto_route 
            })
            # Draw route on map
            map_path = self.map_widget.set_path(map_coords_list, color=route_color_for_auto_route, width=3)
            self.map_paths.append(map_path)
            routes_created_count += 1
            
        messagebox.showinfo("Rutas Automáticas", f"Se crearon {routes_created_count} rutas automáticas.")

    def on_color_change(self, event):
        """
        Handles the `<<ComboboxSelected>>` event for the route color combobox.

        If any pins are currently selected in the UI (via their checkbuttons),
        this method automatically triggers the creation of a route using these
        selected pins. The newly selected color from the combobox is used for this route.
        After the route is created, all pins are deselected. This provides a
        quick way to create a route with a specific color if pins are already chosen.

        Args:
            event: The Tkinter event object associated with the combobox selection.
                   This argument is passed by the event binding but is not directly
                   used in this method's logic.
        """
        # Check if any pins are currently selected
        selected_pins = [pin for pin in self.pins_data if pin["tk_var"].get()]
        if selected_pins: # Only proceed if there's a selection
            # Create a route using the current selection and the newly chosen color (which is already set in the combobox)
            self.create_route_from_selection()
            # Deselect all pins after the route is created for convenience
            self.deselect_all_pins()
        self._apply_theme() # Re-apply theme as combobox interaction might affect styles


if __name__ == "__main__":
    # This block runs when the script is executed directly.
    app = KMZRouteApp() # Create an instance of the application
    app.mainloop()      # Start the Tkinter event loop to display the UI and handle events
