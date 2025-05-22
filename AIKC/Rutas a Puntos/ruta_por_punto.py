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


class KMZRouteApp(tkinter.Tk):
    def __init__(self):
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

        self._setup_ui()
        self.map_widget.set_position(-25.2637, -57.5759) # Asunción, Paraguay
        self.map_widget.set_zoom(5)

    def _setup_ui(self):
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
        self.route_color_combo = ttk.Combobox(route_controls_frame, values=["cyan", "rojo", "verde", "azul"], state="readonly")
        self.route_color_combo.current(1)  # Valor por defecto: "rojo"
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
        canvas_width = event.width
        self.pins_canvas.itemconfig(self.pins_canvas_window, width=canvas_width)

    def _clear_pin_list_ui(self):
        for widget in self.pins_list_frame.winfo_children():
            widget.destroy()

    def _clear_map_markers(self):
        for marker in self.map_markers:
            marker.delete()
        self.map_markers = []

    def _clear_map_paths(self):
        for path in self.map_paths:
            path.delete()
        self.map_paths = []

    def clear_map_and_data(self):
        self._clear_pin_list_ui()
        self._clear_map_markers()
        self._clear_map_paths()
        self.pins_data = []
        self.routes_data = []
        self.route_name_entry.delete(0, tkinter.END)
        self.map_widget.set_zoom(5)
        messagebox.showinfo("Limpieza Completa", "Se han eliminado todos los pines y rutas del mapa y la aplicación.")

    def load_kmz_file(self):
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
            self._extract_placemarks_from_lxml_tree(xml_root)

            if not self.pins_data:
                messagebox.showinfo("Información", "No se encontraron pines (Placemarks con Puntos) en el archivo KMZ.")
            else:
                self._populate_pin_list_ui() 
                self._zoom_to_pins() 
                messagebox.showinfo("KMZ Cargado", f"Se cargaron {len(self.pins_data)} pines desde {self.current_source}.")

        except Exception as e:
            messagebox.showerror("Error al Cargar KMZ", f"Ocurrió un error: {e}")
            import traceback
            print(traceback.format_exc()) 

    def _extract_placemarks_from_lxml_tree(self, xml_element):
        """
        Extrae recursivamente Placemarks con geometría de Punto de un árbol lxml.
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
                
                if point_element is not None:
                    coordinates_element = point_element.find(f"{KML_NS}coordinates")
                    if coordinates_element is not None and coordinates_element.text:
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
                            pass
                    else:
                        pass
                else:
                    pass

    def _populate_pin_list_ui(self):
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
        # Si shift está presionado, se seleccionarán todos los pines entre el último clic y el actual.
        SHIFT_MASK = 0x0001
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
        # Alterna el estado y actualiza el color del marcador.
        pin_info["tk_var"].set(not pin_info["tk_var"].get())
        self.update_marker_color(pin_info)

    def _zoom_to_pins(self):
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
        selected_pins = [pin for pin in self.pins_data if pin["tk_var"].get()]

        if len(selected_pins) < 2:
            messagebox.showwarning("Selección Insuficiente", "Seleccione al menos dos pines para crear una ruta.")
            return

        route_name = self.route_name_entry.get().strip()
        if not route_name: 
            route_name = f"Ruta-{len(self.routes_data) + 1}"
            self.route_name_entry.insert(0, route_name)
        # Usar el valor del combobox para el color y convertirlo a un nombre reconocible
        route_color = self.route_color_combo.get().strip() or "rojo"
        color_mapping = {"rojo": "red", "verde": "green", "azul": "blue", "cyan": "cyan"}
        route_color_mapped = color_mapping.get(route_color.lower(), route_color)
        
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

        kml_output = simplekml.Kml(name="Rutas Generadas")  # ...existing code...
        # Nuevo mapeo de colores para KML (aabbggrr)
        kml_color_mapping = {"red": "ff0000ff", "green": "ff00ff00", "blue": "ffff0000", "cyan": "ffffff00"}
        for route_info in self.routes_data:
            route_name = route_info["name"]
            coords = route_info["kml_coords"]  # ...existing code...
            linestring = kml_output.newlinestring(name=route_name, coords=coords)
            route_color = route_info.get("color", "red")
            linestring.style.linestyle.color = kml_color_mapping.get(route_color.lower(), "ff0000ff")
            linestring.style.linestyle.width = 3
        try:
            kml_output.save(filepath)  # ...existing code...
            messagebox.showinfo("Guardado Exitoso", f"Rutas guardadas en '{os.path.basename(filepath)}' usando SimpleKML.")
        except Exception as e:
            messagebox.showerror("Error al Guardar con SimpleKML", f"No se pudo guardar el archivo KML: {e}")

    def select_all_pins(self):
        """Selecciona todos los pines en la UI marcando sus checkbuttons."""
        for pin in self.pins_data:
            pin["tk_var"].set(True)

    def deselect_all_pins(self):
        """Deselecciona todos los pines en la UI desmarcando sus checkbuttons."""
        for pin in self.pins_data:
            pin["tk_var"].set(False)

    def update_ordering(self):
        # Recorrer pines para asignar o limpiar orden según estado.
        selected = []
        for pin in self.pins_data:
            if pin["tk_var"].get():
                if "select_order" not in pin or pin["select_order"] is None:
                    pin["select_order"] = self.order_counter
                    self.order_counter += 1
                selected.append(pin)
            else:
                pin["select_order"] = None
        selected.sort(key=lambda p: p["select_order"])
        for pin in self.pins_data:
            base_name = pin["name"]
            order_prefix = ""
            if pin["tk_var"].get() and pin["select_order"] is not None:
                pos = next((j for j, p in enumerate(selected) if p is pin), None)
                if pos is not None:
                    order_prefix = f"{pos+1}. "
            # Verifica que el widget exista antes de actualizar su texto
            if "checkbox_widget" in pin and pin["checkbox_widget"].winfo_exists():
                pin["checkbox_widget"].config(text=order_prefix + base_name)
            # NUEVO: actualizar el color de su marcador
            self.update_marker_color(pin)

    def update_marker_color(self, pin):
        # Actualiza el color del marcador según si el pin está seleccionado.
        new_color = "green" if pin["tk_var"].get() else "red"
        if "map_marker" in pin and pin["map_marker"]:
            pin["map_marker"].delete()
        new_marker = self.map_widget.set_marker(
            pin["coords_map"][0],
            pin["coords_map"][1],
            text=pin["name"],
            marker_color_circle=new_color,  # Se usa marker_color_circle
            command=lambda m, p=pin: self._on_marker_click(p)
        )
        pin["map_marker"] = new_marker

    # NUEVO: método para agrupar llamadas a update_ordering
    def schedule_update_ordering(self):
        if self.update_ordering_id is not None:
            self.after_cancel(self.update_ordering_id)
        self.update_ordering_id = self.after(100, self.update_ordering)

    # NUEVO: método para crear rutas automáticas agrupando los pines por fuente (KMZ)
    def create_routes_from_all(self):
        groups = {}
        for pin in self.pins_data:
            src = pin.get("source", "Sin Fuente")
            groups.setdefault(src, []).append(pin)
        count = 0
        for src, pins in groups.items():
            if len(pins) < 2:
                continue
            route_name = f"Ruta {src}"
            route_kml_coords = [pin["coords_original"] for pin in pins]
            map_coords_list = [pin["coords_map"] for pin in pins]
            # Se puede elegir un color predeterminado o calcular uno; se usa "red" por defecto
            route_color = "red"
            self.routes_data.append({
                "name": route_name,
                "kml_coords": route_kml_coords,
                "color": route_color
            })
            map_path = self.map_widget.set_path(map_coords_list, color=route_color, width=3)
            self.map_paths.append(map_path)
            count += 1
        messagebox.showinfo("Rutas Automáticas", f"Se crearon {count} rutas automáticas.")

    # NUEVO: método para disparar creación de ruta automáticamente al cambiar de color
    def on_color_change(self, event):
        selected_pins = [pin for pin in self.pins_data if pin["tk_var"].get()]
        if selected_pins:
            self.create_route_from_selection()
            self.deselect_all_pins()


if __name__ == "__main__":
    app = KMZRouteApp()
    app.mainloop()
