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

# Namespaces comunes en KML
KML_NS = "{http://www.opengis.net/kml/2.2}"
GX_NS = "{http://www.google.com/kml/ext/2.2}"
ATOM_NS = "{http://www.w3.org/2005/Atom}"
NS_MAP = {
    'kml': 'http://www.opengis.net/kml/2.2',
    'gx': 'http://www.google.com/kml/ext/2.2',
    'atom': 'http://www.w3.org/2005/Atom'
}

# Color Constants
COLOR_RED_NAME = "rojo" # User-facing name
COLOR_GREEN_NAME = "verde" # User-facing name
COLOR_BLUE_NAME = "azul" # User-facing name
COLOR_CYAN_NAME = "cyan" # User-facing name

COLOR_RED = "red"
COLOR_GREEN = "green"
COLOR_BLUE = "blue"
COLOR_CYAN = "cyan"

DEFAULT_ROUTE_COLOR_UI_NAME = COLOR_RED_NAME 
DEFAULT_ROUTE_COLOR_INTERNAL = COLOR_RED
DEFAULT_MARKER_COLOR = COLOR_RED
SELECTED_MARKER_COLOR = COLOR_GREEN

KML_COLOR_RED = "ff0000ff"
KML_COLOR_GREEN = "ff00ff00"
KML_COLOR_BLUE = "ffff0000"
KML_COLOR_CYAN = "ffffff00"
DEFAULT_KML_COLOR = KML_COLOR_RED


class KMZRouteApp(tkinter.Tk):
    def __init__(self):
        """
        Initializes the KMZRouteApp application.

        Sets up the main window, initializes data structures for pins and routes,
        configures map widget default position and zoom, and calls the UI setup method.
        """
        super().__init__()
        self.title("Visor KMZ con LXML y SimpleKML")
        self.geometry("1200x800")

        self.pins_data = []
        self.routes_data = []
        self.map_markers = []
        self.map_paths = []
        self.last_selected_index = None  # índice del último pin clickeado
        self.order_counter = 1  # NUEVO: contador para asignar orden de selección
        self.update_ordering_id = None  # NUEVO: ID para agrupar actualizaciones de orden
        self.extraction_error_count = 0 # For counting errors during placemark extraction

        self._setup_ui()
        self.map_widget.set_position(-25.2637, -57.5759) # Asunción, Paraguay
        self.map_widget.set_zoom(5)

    def _setup_ui(self):
        """
        Sets up the user interface of the application.

        Creates and arranges all the UI elements like frames, buttons, lists,
        and the map widget.
        """
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(expand=True, fill="both")

        paned_window = ttk.PanedWindow(main_frame, orient="horizontal")
        paned_window.pack(expand=True, fill="both")

        left_panel = ttk.Frame(paned_window, width=350, padding="5")
        left_panel.pack_propagate(False)
        paned_window.add(left_panel, weight=1)

        load_button = ttk.Button(left_panel, text="Cargar Archivo KMZ", command=self.load_kmz_file)
        load_button.pack(pady=10, padx=5, fill="x")

        ttk.Separator(left_panel, orient="horizontal").pack(fill="x", pady=5)

        pins_list_frame_container = ttk.LabelFrame(left_panel, text="Pines Disponibles", padding="5")
        pins_list_frame_container.pack(expand=True, fill="both", pady=5, padx=5)

        self.pins_canvas = tkinter.Canvas(pins_list_frame_container, borderwidth=0)
        self.pins_list_frame = ttk.Frame(self.pins_canvas)
        scrollbar = ttk.Scrollbar(pins_list_frame_container, orient="vertical", command=self.pins_canvas.yview)
        self.pins_canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.pins_canvas.pack(side="left", fill="both", expand=True)
        self.pins_canvas_window = self.pins_canvas.create_window((0, 0), window=self.pins_list_frame, anchor="nw")

        self.pins_list_frame.bind("<Configure>", lambda e: self.pins_canvas.configure(scrollregion=self.pins_canvas.bbox("all")))
        self.pins_canvas.bind('<Configure>', self._on_canvas_configure)

        route_controls_frame = ttk.LabelFrame(left_panel, text="Crear Ruta", padding="5")
        route_controls_frame.pack(fill="x", pady=10, padx=5)

        ttk.Label(route_controls_frame, text="Nombre de la Ruta:").pack(anchor="w", padx=5)
        self.route_name_entry = ttk.Entry(route_controls_frame)
        self.route_name_entry.pack(fill="x", padx=5, pady=(0,5))
        # Reemplazar el campo de texto por un Combobox para seleccionar el color.
        ttk.Label(route_controls_frame, text="Color de la Ruta:").pack(anchor="w", padx=5)
        self.route_color_combo_values = [COLOR_CYAN_NAME, COLOR_RED_NAME, COLOR_GREEN_NAME, COLOR_BLUE_NAME]
        self.route_color_combo = ttk.Combobox(route_controls_frame, values=self.route_color_combo_values, state="readonly")
        try:
            default_color_index = self.route_color_combo_values.index(DEFAULT_ROUTE_COLOR_UI_NAME)
            self.route_color_combo.current(default_color_index)
        except ValueError:
            self.route_color_combo.current(0) # Default to first if not found (e.g. cyan)
        self.route_color_combo.pack(fill="x", padx=5, pady=(0,5))
        # NUEVO: cuando se selecciona un color se invoca la creación automática
        self.route_color_combo.bind("<<ComboboxSelected>>", self.on_color_change)
        create_route_button = ttk.Button(route_controls_frame, text="Crear Ruta con Pines Seleccionados", command=self.create_route_from_selection)
        create_route_button.pack(pady=5, fill="x", padx=5)
        
        # NUEVO: botón para crear rutas automáticas agrupadas por fuente
        auto_routes_button = ttk.Button(route_controls_frame, text="Crear Rutas Automáticas", command=self.create_routes_from_all)
        auto_routes_button.pack(pady=5, fill="x", padx=5)

        # Nuevos botones para selección múltiple
        select_buttons_frame = ttk.Frame(route_controls_frame)
        select_buttons_frame.pack(fill="x", pady=5, padx=5)
        
        select_all_button = ttk.Button(select_buttons_frame, text="Seleccionar Todos", command=self.select_all_pins)
        select_all_button.pack(side="left", expand=True, fill="x", padx=(0,2))
        deselect_all_button = ttk.Button(select_buttons_frame, text="Deseleccionar Todos", command=self.deselect_all_pins)
        deselect_all_button.pack(side="left", expand=True, fill="x", padx=(2,0))

        save_routes_button = ttk.Button(left_panel, text="Guardar Rutas Generadas (KML con SimpleKML)", command=self.save_routes_to_kml)
        save_routes_button.pack(pady=10, padx=5, fill="x")
        
        clear_map_button = ttk.Button(left_panel, text="Limpiar Mapa (Pines y Rutas)", command=self.clear_map_and_data)
        clear_map_button.pack(pady=5, padx=5, fill="x")

        map_frame = ttk.Frame(paned_window, padding="5")
        paned_window.add(map_frame, weight=3)

        self.map_widget = tkintermapview.TkinterMapView(map_frame, corner_radius=0)
        self.map_widget.pack(expand=True, fill="both")

    def _on_canvas_configure(self, event):
        """
        Handles the configure event for the pins canvas.

        Adjusts the width of the frame inside the canvas to match the canvas width,
        ensuring the scrollbar behaves correctly.

        Args:
            event: The event object containing details about the configure event.
        """
        canvas_width = event.width
        self.pins_canvas.itemconfig(self.pins_canvas_window, width=canvas_width)

    def _clear_pin_list_ui(self):
        """Clears all widgets from the pin list UI frame."""
        for widget in self.pins_list_frame.winfo_children():
            widget.destroy()

    def _clear_map_markers(self):
        """Removes all markers from the map widget and clears the internal list."""
        for marker in self.map_markers:
            marker.delete()
        self.map_markers = []

    def _clear_map_paths(self):
        """Removes all paths from the map widget and clears the internal list."""
        for path in self.map_paths:
            path.delete()
        self.map_paths = []

    def clear_map_and_data(self):
        """
        Clears all loaded data, UI elements, and map features.

        Resets the application to its initial state by clearing the pin list,
        map markers, map paths, internal data storage for pins and routes,
        and resets the route name entry field. Also, shows an info message.
        """
        self._clear_pin_list_ui()
        self._clear_map_markers()
        self._clear_map_paths()
        self.pins_data = []
        self.routes_data = []
        self.route_name_entry.delete(0, tkinter.END)
        self.map_widget.set_zoom(5)
        messagebox.showinfo("Limpieza Completa", "Se han eliminado todos los pines y rutas del mapa y la aplicación.")

    def load_kmz_file(self):
        """
        Loads pins from a KMZ file selected by the user.

        Opens a file dialog for KMZ selection, then parses the KML content
        within the KMZ to extract placemark data. Updates the UI with loaded pins.
        Handles potential errors during file operations or parsing.
        """
        filepath = filedialog.askopenfilename(
            title="Seleccionar Archivo KMZ",
            filetypes=(("Archivos KMZ", "*.kmz"), ("Todos los archivos", "*.*"))
        )
        if not filepath:
            return

        self.clear_map_and_data() 
        # NUEVO: Registrar la fuente actual
        self.current_source = os.path.basename(filepath)

        try:
            with zipfile.ZipFile(filepath, 'r') as kmz:
                kml_filename = None
                for name in kmz.namelist():
                    if name.lower().endswith('.kml'):
                        kml_filename = name
                        break
                
                if not kml_filename:
                    messagebox.showerror("Error en KMZ", "No se encontró un archivo KML dentro del KMZ.")
                    return

                kml_bytes = kmz.read(kml_filename)
                
            parser = etree.XMLParser(resolve_entities=False, strip_cdata=False, remove_comments=True)
            xml_root = etree.fromstring(kml_bytes, parser=parser)
            
            self.pins_data = []
            self.extraction_error_count = 0 # Reset error count for this load
            self._extract_placemarks_from_lxml_tree(xml_root)

            num_loaded = len(self.pins_data)
            num_skipped = self.extraction_error_count
            source_name = self.current_source

            if num_loaded > 0:
                success_msg = f"Se cargaron {num_loaded} pines desde {source_name}."
                if num_skipped > 0:
                    skipped_msg = f" Se omitieron {num_skipped} pines debido a errores en el formato de coordenadas."
                    messagebox.showinfo("KMZ Cargado Parcialmente", success_msg + skipped_msg)
                else:
                    messagebox.showinfo("KMZ Cargado", success_msg)
                self._populate_pin_list_ui() 
                self._zoom_to_pins()
            else: # No pins loaded
                if num_skipped > 0:
                    messagebox.showwarning("Error de Carga de Pines", f"No se cargaron pines desde {source_name}. Se omitieron {num_skipped} pines debido a errores en el formato de coordenadas.")
                else:
                    messagebox.showinfo("Información", f"No se encontraron pines (Placemarks con Puntos) en el archivo KMZ '{source_name}'.")

        except Exception as e:
            messagebox.showerror("Error al Cargar KMZ", f"Ocurrió un error: {e}")
            import traceback
            print(traceback.format_exc()) 

    def _extract_placemarks_from_lxml_tree(self, xml_element):
        """
        Recursively extracts Placemarks with Point geometry from an lxml tree.

        Iterates through XML elements, looking for KML Documents, Folders,
        and Placemarks. For Placemarks with a Point, it extracts the name
        and coordinates, storing them in `self.pins_data`.

        Args:
            xml_element: The root lxml element to start parsing from.
        """
        for child in xml_element:
            if isinstance(child, etree._Comment):
                continue
            
            if child.tag == f"{KML_NS}Document" or child.tag == f"{KML_NS}Folder":
                self._extract_placemarks_from_lxml_tree(child)
            
            elif child.tag == f"{KML_NS}Placemark":
                placemark_name_element = child.find(f"{KML_NS}name")
                placemark_name = placemark_name_element.text if placemark_name_element is not None and placemark_name_element.text else "Pin sin nombre"
                
                point_element = child.find(f".//{KML_NS}Point")
                
                if point_element is None:
                    continue

                coordinates_element = point_element.find(f"{KML_NS}coordinates")
                if coordinates_element is None or not coordinates_element.text:
                    continue

                coords_str = coordinates_element.text.strip()
                try:
                    lon_str, lat_str, *alt_str = coords_str.split(',')
                    lon = float(lon_str)
                    lat = float(lat_str)
                    alt = float(alt_str[0]) if alt_str else 0.0

                    pin_info = {
                        "name": placemark_name,
                        "coords_original": (lon, lat, alt),
                        "coords_map": (lat, lon),
                        "tk_var": tkinter.BooleanVar(value=False),
                        # NUEVO: se guarda la fuente del KMZ
                        "source": getattr(self, "current_source", "Sin Fuente")
                    }
                    self.pins_data.append(pin_info)
                except ValueError:
                    # If coordinates are malformed, skip this placemark
                    self.extraction_error_count += 1
                    pass # Continue to the next placemark/child

    def _populate_pin_list_ui(self):
        """
        Populates the UI list with available pins and places markers on the map.

        Clears any existing pins from the UI and map. Then, for each pin in
        `self.pins_data`, it creates a checkbutton in the list and a marker
        on the map. Binds click events for selection and ordering.
        """
        self._clear_pin_list_ui() 
        self._clear_map_markers()

        for i, pin in enumerate(self.pins_data):
            cb = ttk.Checkbutton(self.pins_list_frame, text=pin["name"], variable=pin["tk_var"])
            cb.pack(anchor="w", fill="x", padx=5)
            pin["checkbox_widget"] = cb 
            # Vincular para selección con shift
            cb.bind("<Button-1>", lambda event, index=i: self.on_checkbutton_click(event, index))
            # Se utiliza la actualización agrupada en lugar de actualizar inmediatamente
            pin["tk_var"].trace_add("write", lambda *args: self.schedule_update_ordering())

            marker = self.map_widget.set_marker(
                pin["coords_map"][0],  
                pin["coords_map"][1],  
                text=pin["name"],
                command=lambda m, p=pin: self._on_marker_click(p) 
            )
            self.map_markers.append(marker)
            pin["map_marker"] = marker 
        
        self.pins_list_frame.update_idletasks()
        self.pins_canvas.config(scrollregion=self.pins_canvas.bbox("all"))

    def on_checkbutton_click(self, event, index):
        """
        Handles click events on pin checkbuttons for selection.

        Supports range selection using the Shift key. If Shift is pressed,
        all pins between the last selected pin and the currently clicked pin
        will have their selection state set to that of the currently clicked pin.
        Updates `self.last_selected_index`.

        Args:
            event: The Tkinter event object.
            index: The index of the clicked pin in `self.pins_data`.

        Returns:
            "break" if Shift selection was performed, to prevent default behavior.
        """
        # Si shift está presionado, se seleccionarán todos los pines entre el último clic y el actual.
        SHIFT_MASK = 0x0001 # Standard Tkinter mask for Shift key
        if event.state & SHIFT_MASK and self.last_selected_index is not None:
            start = min(self.last_selected_index, index)
            end = max(self.last_selected_index, index)
            # Determinar la acción: si el pin actual está sin seleccionar, se selecciona; de lo contrario se deselecciona.
            new_state = not self.pins_data[index]["tk_var"].get()
            for i in range(start, end+1):
                self.pins_data[i]["tk_var"].set(new_state)
            # Actualizar el índice final y evitar acción normal
            self.last_selected_index = index
            return "break"
        else:
            # Actualiza el índice sin acción especial.
            self.last_selected_index = index

    def _on_marker_click(self, pin_info):
        """
        Handles click events on map markers.

        Toggles the selection state of the associated pin (`pin_info["tk_var"]`)
        and updates the marker's color to reflect the new state.

        Args:
            pin_info: The dictionary containing information about the clicked pin.
        """
        # Alterna el estado y actualiza el color del marcador.
        pin_info["tk_var"].set(not pin_info["tk_var"].get())
        self.update_marker_color(pin_info)

    def _zoom_to_pins(self):
        """
        Adjusts the map's viewport to encompass all loaded pins.

        If there's only one pin, it centers the map on that pin with a fixed zoom level.
        If there are multiple pins, it fits the map to the bounding box of all pins.
        Does nothing if no pins are loaded.
        """
        if not self.map_markers:
            return
        
        if len(self.map_markers) == 1: 
            marker = self.map_markers[0]
            self.map_widget.set_position(marker.position[0], marker.position[1])
            self.map_widget.set_zoom(15) 
            return

        marker_positions = [marker.position for marker in self.map_markers]
        if marker_positions:
            lats = [pos[0] for pos in marker_positions]
            lons = [pos[1] for pos in marker_positions]
            top_left = (max(lats), min(lons))
            bottom_right = (min(lats), max(lons))
            self.map_widget.fit_bounding_box(top_left, bottom_right)

    def create_route_from_selection(self):
        """
        Creates a new route from the currently selected pins.

        Requires at least two pins to be selected. The route name is taken
        from an entry field or generated if empty. The route color is taken
        from a combobox. The new route is added to `self.routes_data` and
        drawn on the map.
        """
        selected_pins = [pin for pin in self.pins_data if pin["tk_var"].get()]

        if len(selected_pins) < 2:
            messagebox.showwarning("Selección Insuficiente", "Seleccione al menos dos pines para crear una ruta.")
            return

        route_name = self.route_name_entry.get().strip()
        if not route_name: 
            route_name = f"Ruta-{len(self.routes_data) + 1}"
            self.route_name_entry.insert(0, route_name)

        route_color_ui_name = self.route_color_combo.get().strip() or DEFAULT_ROUTE_COLOR_UI_NAME
        
        # Map user-facing names to internal color names
        ui_to_internal_color_mapping = {
            COLOR_CYAN_NAME: COLOR_CYAN,
            COLOR_RED_NAME: COLOR_RED,
            COLOR_GREEN_NAME: COLOR_GREEN,
            COLOR_BLUE_NAME: COLOR_BLUE,
        }
        route_color_mapped = ui_to_internal_color_mapping.get(route_color_ui_name, DEFAULT_ROUTE_COLOR_INTERNAL)
        
        # Coordenadas para la ruta (lon, lat, alt) y para el path (lat, lon)
        route_kml_coords = [pin["coords_original"] for pin in selected_pins]
        map_coords_list = [pin["coords_map"] for pin in selected_pins]

        self.routes_data.append({
            "name": route_name,
            "kml_coords": route_kml_coords,  # ...existing code...
            "color": route_color_mapped
        })

        map_path = self.map_widget.set_path(map_coords_list, color=route_color_mapped, width=3)
        self.map_paths.append(map_path) 

        messagebox.showinfo("Ruta Creada", f"Ruta '{route_name}' creada con {len(selected_pins)} puntos y añadida al mapa.")
        # Limpia el campo para que cada ruta creada conserve el nombre asignado en su creación
        self.route_name_entry.delete(0, tkinter.END)

    def save_routes_to_kml(self):
        """
        Saves all created routes to a KML file using simplekml.

        Prompts the user for a save location. If routes exist, it generates
        a KML structure with linestrings for each route, applying specified colors
        based on internal names mapped to KML color codes.
        Handles potential errors during file saving.
        """
        if not self.routes_data:
            messagebox.showinfo("Sin Rutas", "No hay rutas creadas para guardar.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Guardar Rutas como KML (con SimpleKML)",
            defaultextension=".kml",
            filetypes=(("Archivos KML", "*.kml"), ("Todos los archivos", "*.*"))
        )
        if not filepath:
            return

        kml_output = simplekml.Kml(name="Rutas Generadas")
        # Map internal color names to KML color codes
        internal_to_kml_color_mapping = {
            COLOR_RED: KML_COLOR_RED,
            COLOR_GREEN: KML_COLOR_GREEN,
            COLOR_BLUE: KML_COLOR_BLUE,
            COLOR_CYAN: KML_COLOR_CYAN,
        }
        for route_info in self.routes_data:
            route_name = route_info["name"]
            coords = route_info["kml_coords"]
            linestring = kml_output.newlinestring(name=route_name, coords=coords)
            route_color_internal_name = route_info.get("color", DEFAULT_ROUTE_COLOR_INTERNAL)
            linestring.style.linestyle.color = internal_to_kml_color_mapping.get(route_color_internal_name, DEFAULT_KML_COLOR)
            linestring.style.linestyle.width = 3
        try:
            kml_output.save(filepath)
            messagebox.showinfo("Guardado Exitoso", f"Rutas guardadas en '{os.path.basename(filepath)}' usando SimpleKML.")
        except Exception as e:
            messagebox.showerror("Error al Guardar con SimpleKML", f"No se pudo guardar el archivo KML: {e}")

    def select_all_pins(self):
        """
        Selects all pins in the UI.

        Sets the `tk_var` (Tkinter BooleanVar) of each pin to True, which
        checks the corresponding checkbutton and triggers associated updates
        (like reordering and marker color change).
        """
        for pin in self.pins_data:
            pin["tk_var"].set(True)

    def deselect_all_pins(self):
        """
        Deselects all pins in the UI.

        Sets the `tk_var` (Tkinter BooleanVar) of each pin to False, which
        unchecks the corresponding checkbutton and triggers associated updates
        (like reordering and marker color change).
        """
        for pin in self.pins_data:
            pin["tk_var"].set(False)

    def update_ordering(self):
        """
        Updates the displayed order of selected pins and their marker colors.

        Assigns a selection order number to newly selected pins based on `self.order_counter`.
        Updates the text of pin checkbuttons in the UI to show their selection order (e.g., "1. Pin Name").
        Resets the text to the base name if a pin is deselected or its order is cleared.
        Updates marker colors for all pins to reflect their current selection status.
        Resets `self.order_counter` to 1 if no pins are selected to ensure fresh ordering for the next selection.
        """
        selected = []
        for pin in self.pins_data:
            if pin["tk_var"].get():
                if "select_order" not in pin or pin["select_order"] is None:
                    pin["select_order"] = self.order_counter
                    self.order_counter += 1
                selected.append(pin)
            else:
                pin["select_order"] = None # Clear order if not selected

        # Reset order_counter if nothing is selected, so next selection starts from 1
        if not any(p["tk_var"].get() for p in self.pins_data):
            self.order_counter = 1
        
        # Sort selected pins by their assigned order number
        selected.sort(key=lambda p: p["select_order"] if p["select_order"] is not None else float('inf'))

        # Update checkbox text for selected pins with their order
        for i, pin_selected in enumerate(selected):
             if pin_selected["tk_var"].get() and pin_selected["select_order"] is not None:
                base_name = pin_selected["name"]
                order_prefix = f"{i+1}. "
                if "checkbox_widget" in pin_selected and pin_selected["checkbox_widget"].winfo_exists():
                    pin_selected["checkbox_widget"].config(text=order_prefix + base_name)

        # Update all pins (reset text for deselected, update marker color for all)
        for pin in self.pins_data:
            if not pin["tk_var"].get() or pin["select_order"] is None: # If not selected or order cleared
                if "checkbox_widget" in pin and pin["checkbox_widget"].winfo_exists():
                     pin["checkbox_widget"].config(text=pin["name"]) # Reset to base name
            self.update_marker_color(pin)

    def update_marker_color(self, pin):
        """
        Updates the color of a map marker based on its selection state.

        Sets the marker color to `SELECTED_MARKER_COLOR` if selected,
        or `DEFAULT_MARKER_COLOR` otherwise. Recreates the marker on the map
        to apply the color change and re-binds its click command.

        Args:
            pin: The pin dictionary, which contains its selection state (`tk_var`)
                 and map marker object (`map_marker`).
        """
        new_color = SELECTED_MARKER_COLOR if pin["tk_var"].get() else DEFAULT_MARKER_COLOR
        if "map_marker" in pin and pin["map_marker"]:
            pin["map_marker"].delete() # Delete old marker
        
        # Create new marker with updated color
        new_marker = self.map_widget.set_marker(
            pin["coords_map"][0],
            pin["coords_map"][1],
            text=pin["name"],
            marker_color_circle=new_color, 
            command=lambda m, p=pin: self._on_marker_click(p) # Re-bind command
        )
        pin["map_marker"] = new_marker # Store new marker

    def schedule_update_ordering(self):
        """
        Schedules a delayed call to `update_ordering`.

        This is used to group multiple rapid changes (e.g., from checkbutton
        traces or multiple selections) into a single UI update, improving performance
        and responsiveness. Cancels any previously scheduled update before scheduling
        a new one. The delay is 100ms.
        """
        if self.update_ordering_id is not None:
            self.after_cancel(self.update_ordering_id)
        self.update_ordering_id = self.after(100, self.update_ordering) # 100ms delay

    def create_routes_from_all(self):
        """
        Automatically creates routes by grouping all loaded pins by their source KMZ file.

        For each unique source file identified in the pin data, if there are at
        least two pins from that source, a route is created. Routes are given a
        default name based on the source (e.g., "Ruta example.kmz") and use
        `DEFAULT_ROUTE_COLOR_INTERNAL`. Displays a message with the number of routes created.
        """
        groups = {}
        for pin in self.pins_data:
            src = pin.get("source", "Sin Fuente") # Default if source is missing
            groups.setdefault(src, []).append(pin)
        
        routes_created_count = 0 # Changed from 'count'
        for src, pins_in_group in groups.items(): # Changed from 'pins'
            if len(pins_in_group) < 2: # Changed from 'pins'
                continue # Need at least two pins for a route
            
            route_name = f"Ruta {src}"
            # Pins are used in the order they appear in pins_in_group for this source.
            # If a specific order is needed, pins_in_group should be sorted here.
            route_kml_coords = [p["coords_original"] for p in pins_in_group] # Changed from 'pin' to 'p'
            map_coords_list = [p["coords_map"] for p in pins_in_group] # Changed from 'p["coords_map"] for p in pins_in_group'
            
            route_color_for_auto_route = DEFAULT_ROUTE_COLOR_INTERNAL 
            
            self.routes_data.append({
                "name": route_name,
                "kml_coords": route_kml_coords,
                "color": route_color_for_auto_route 
            })
            map_path = self.map_widget.set_path(map_coords_list, color=route_color_for_auto_route, width=3)
            self.map_paths.append(map_path)
            routes_created_count += 1
            
        messagebox.showinfo("Rutas Automáticas", f"Se crearon {routes_created_count} rutas automáticas.")

    def on_color_change(self, event):
        """
        Handles the event triggered when the route color is changed in the combobox.

        If any pins are currently selected in the UI, this method automatically
        creates a route using these selected pins and the newly chosen color from
        the combobox. After route creation, all pins are deselected.

        Args:
            event: The Tkinter event object (passed by the `<<ComboboxSelected>>`
                   binding, not directly used in the method body).
        """
        selected_pins = [pin for pin in self.pins_data if pin["tk_var"].get()]
        if selected_pins: # Only create route if pins are selected
            self.create_route_from_selection()
            self.deselect_all_pins()


if __name__ == "__main__":
    app = KMZRouteApp()
    app.mainloop()
