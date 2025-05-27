import unittest
from unittest.mock import patch, MagicMock, mock_open, call
import sys
import os

# Mock modules before importing the application
MOCK_MODULES = {
    'tkinter': MagicMock(),
    'tkinter.ttk': MagicMock(),
    'tkinter.filedialog': MagicMock(),
    'tkinter.messagebox': MagicMock(),
    'tkintermapview': MagicMock(),
    'simplekml': MagicMock(),
    'lxml': MagicMock(),
    'lxml.etree': MagicMock(),
}

# Helper to create mock lxml elements
def create_mock_element(tag, text=None, children=None, attrib=None):
    el = MagicMock()
    el.tag = tag
    el.text = text
    el.attrib = attrib if attrib else {}
    
    # Children can be a list of other mock elements or a callable that returns a list
    _children = children if children else []
    
    # Configure iteration over children
    el.__iter__.return_value = iter(_children)
    
    # Configure find and findall
    # This is a simplified find/findall, may need to be more sophisticated for complex queries
    def find_mock(path):
        # Basic support for direct children, e.g., "{ns}name"
        # And for descendant search, e.g., ".//{ns}Point"
        is_descendant_search = path.startswith('.//')
        search_tag = path.split('}')[-1] if '}' in path else path
        
        if is_descendant_search:
            queue = list(_children)
            while queue:
                current_child = queue.pop(0)
                # Simplistic tag matching (ignores namespace for descendant search for now)
                if search_tag in current_child.tag:
                    return current_child
                queue.extend(list(current_child.__iter__()))
            return None
        else:
            for child_el in _children:
                if child_el.tag == path: # Path here should include namespace
                    return child_el
            return None
            
    def findall_mock(path):
        # Simplified findall, returns all direct children matching the tag
        # Does not handle complex XPath, only direct child tag matching
        found_elements = []
        for child_el in _children:
            if child_el.tag == path: # Path here should include namespace
                found_elements.append(child_el)
        return found_elements

    el.find = MagicMock(side_effect=find_mock)
    el.findall = MagicMock(side_effect=findall_mock)
    
    # For isinstance checks, like etree._Comment
    if tag == "Comment": # Special handling for comments
        el.__class__ = MOCK_MODULES['lxml.etree']._Comment
        
    return el

# Define KML_NS for convenience in tests
KML_NS_TEST = "{http://www.opengis.net/kml/2.2}"


@patch.dict(sys.modules, MOCK_MODULES)
class TestKMZRouteApp(unittest.TestCase):

    def setUp(self):
        # Import the application class AFTER mocks are in place
        from AIKC.Rutas.a.Puntos.ruta_por_punto import KMZRouteApp, KML_NS, LIGHT_THEME_COLORS, DARK_THEME_COLORS

        # Mock methods that would normally interact with Tkinter's mainloop or UI setup
        with patch.object(KMZRouteApp, '__init__', lambda s: None): # Bypass original __init__
            self.app = KMZRouteApp()

        # Manually initialize attributes that would be set by __init__ or _setup_ui
        self.app.title = MagicMock()
        self.app.geometry = MagicMock()
        self.app.configure = MagicMock() # For root window styling
        self.app.after_cancel = MagicMock()
        self.app.after = MagicMock()
        
        # Data structures
        self.app.pins_data = []
        self.app.routes_data = []
        self.app.map_markers = []
        self.app.map_paths = []
        self.app.last_selected_index = None
        self.app.order_counter = 1
        self.app.update_ordering_id = None
        self.app.extraction_error_count = 0
        self.app.current_source = "test_source.kmz"

        # UI elements that are accessed (mocked)
        self.app.map_widget = MOCK_MODULES['tkintermapview'].TkinterMapView.return_value
        self.app.route_name_entry = MagicMock()
        self.app.route_color_combo = MagicMock()
        self.app.pins_canvas = MagicMock()
        self.app.pins_list_frame = MagicMock() 
        
        # Theme related
        self.app.theme = "light"
        self.app.style = MOCK_MODULES['tkinter.ttk'].Style.return_value
        self.app.LIGHT_THEME_COLORS = LIGHT_THEME_COLORS
        self.app.DARK_THEME_COLORS = DARK_THEME_COLORS
        
        # Constants from the app needed for tests
        self.app.KML_NS = KML_NS # Make sure the app instance uses the real KML_NS

        # Mock _apply_theme as it's called often and can be complex
        self.app._apply_theme = MagicMock()
        
        # Mock tkinter.BooleanVar for pin selection
        self.MockBooleanVar = MOCK_MODULES['tkinter'].BooleanVar


    def test_initial_state(self):
        # For this test, we re-initialize an app instance but let its __init__ run partially
        # to check default initializations before _setup_ui and _apply_theme are called.
        
        # Temporarily unpatch __init__ to test its internal state setup
        with patch.object(sys.modules['AIKC.Rutas.a.Puntos.ruta_por_punto'].KMZRouteApp, '_setup_ui', MagicMock()), \
             patch.object(sys.modules['AIKC.Rutas.a.Puntos.ruta_por_punto'].KMZRouteApp, '_apply_theme', MagicMock()):
            
            # Import here to get the class with unpatched __init__ for this specific test scope
            from AIKC.Rutas.a.Puntos.ruta_por_punto import KMZRouteApp
            app_for_init_test = KMZRouteApp() # This will call the actual __init__ now

        self.assertEqual(len(app_for_init_test.pins_data), 0)
        self.assertEqual(len(app_for_init_test.routes_data), 0)
        self.assertEqual(app_for_init_test.theme, "light")
        self.assertIsNotNone(app_for_init_test.style)
        self.assertEqual(app_for_init_test.order_counter, 1)
        self.assertEqual(app_for_init_test.extraction_error_count, 0)


    def test_extract_placemarks_from_lxml_tree_valid_kml(self):
        placemark1_name = create_mock_element(f"{KML_NS_TEST}name", text="Pin1")
        placemark1_coords = create_mock_element(f"{KML_NS_TEST}coordinates", text="-57.1,-25.1,10")
        placemark1_point = create_mock_element(f"{KML_NS_TEST}Point", children=[placemark1_coords])
        placemark1 = create_mock_element(f"{KML_NS_TEST}Placemark", children=[placemark1_name, placemark1_point])

        placemark2_name = create_mock_element(f"{KML_NS_TEST}name", text="Pin2")
        placemark2_coords = create_mock_element(f"{KML_NS_TEST}coordinates", text="-57.2,-25.2") # No altitude
        placemark2_point = create_mock_element(f"{KML_NS_TEST}Point", children=[placemark2_coords])
        placemark2 = create_mock_element(f"{KML_NS_TEST}Placemark", children=[placemark2_name, placemark2_point])
        
        mock_kml_root = create_mock_element(f"{KML_NS_TEST}kml", children=[placemark1, placemark2])
        
        self.app.current_source = "test.kmz"
        self.app._extract_placemarks_from_lxml_tree(mock_kml_root)

        self.assertEqual(len(self.app.pins_data), 2)
        self.assertEqual(self.app.pins_data[0]["name"], "Pin1")
        self.assertEqual(self.app.pins_data[0]["coords_original"], (-57.1, -25.1, 10.0))
        self.assertEqual(self.app.pins_data[0]["coords_map"], (-25.1, -57.1))
        self.assertEqual(self.app.pins_data[0]["source"], "test.kmz")
        
        self.assertEqual(self.app.pins_data[1]["name"], "Pin2")
        self.assertEqual(self.app.pins_data[1]["coords_original"], (-57.2, -25.2, 0.0)) # Altitude defaults to 0.0
        self.assertEqual(self.app.pins_data[1]["coords_map"], (-25.2, -57.2))
        self.assertEqual(self.app.pins_data[1]["source"], "test.kmz")
        
        self.assertEqual(self.app.extraction_error_count, 0)

    def test_extract_placemarks_from_lxml_tree_malformed_coords(self):
        placemark_name = create_mock_element(f"{KML_NS_TEST}name", text="PinMalformed")
        placemark_coords = create_mock_element(f"{KML_NS_TEST}coordinates", text="not,valid,coords")
        placemark_point = create_mock_element(f"{KML_NS_TEST}Point", children=[placemark_coords])
        placemark_malformed = create_mock_element(f"{KML_NS_TEST}Placemark", children=[placemark_name, placemark_point])

        mock_kml_root = create_mock_element(f"{KML_NS_TEST}kml", children=[placemark_malformed])
        
        self.app._extract_placemarks_from_lxml_tree(mock_kml_root)
        
        self.assertEqual(len(self.app.pins_data), 0)
        self.assertEqual(self.app.extraction_error_count, 1)

    def test_extract_placemarks_from_lxml_tree_no_points(self):
        placemark_name = create_mock_element(f"{KML_NS_TEST}name", text="PinNoPoint")
        # No Point, e.g., a LineString
        linestring_coords = create_mock_element(f"{KML_NS_TEST}coordinates", text="-57.3,-25.3 -57.4,-25.4")
        linestring = create_mock_element(f"{KML_NS_TEST}LineString", children=[linestring_coords])
        placemark_no_point = create_mock_element(f"{KML_NS_TEST}Placemark", children=[placemark_name, linestring])

        mock_kml_root = create_mock_element(f"{KML_NS_TEST}kml", children=[placemark_no_point])
        
        self.app._extract_placemarks_from_lxml_tree(mock_kml_root)
        
        self.assertEqual(len(self.app.pins_data), 0)
        self.assertEqual(self.app.extraction_error_count, 0)

    def test_extract_placemarks_from_lxml_tree_nested_folders(self):
        # Pin1 inside Folder1
        placemark1_name = create_mock_element(f"{KML_NS_TEST}name", text="Pin1Nested")
        placemark1_coords = create_mock_element(f"{KML_NS_TEST}coordinates", text="-57.1,-25.1,10")
        placemark1_point = create_mock_element(f"{KML_NS_TEST}Point", children=[placemark1_coords])
        placemark1 = create_mock_element(f"{KML_NS_TEST}Placemark", children=[placemark1_name, placemark1_point])
        folder1 = create_mock_element(f"{KML_NS_TEST}Folder", children=[placemark1])
        
        # Pin2 at document level
        placemark2_name = create_mock_element(f"{KML_NS_TEST}name", text="Pin2Root")
        placemark2_coords = create_mock_element(f"{KML_NS_TEST}coordinates", text="-57.2,-25.2")
        placemark2_point = create_mock_element(f"{KML_NS_TEST}Point", children=[placemark2_coords])
        placemark2 = create_mock_element(f"{KML_NS_TEST}Placemark", children=[placemark2_name, placemark2_point])

        document = create_mock_element(f"{KML_NS_TEST}Document", children=[folder1, placemark2])
        mock_kml_root = create_mock_element(f"{KML_NS_TEST}kml", children=[document])
        
        self.app.current_source = "nested.kmz"
        self.app._extract_placemarks_from_lxml_tree(mock_kml_root)

        self.assertEqual(len(self.app.pins_data), 2)
        self.assertTrue(any(p["name"] == "Pin1Nested" for p in self.app.pins_data))
        self.assertTrue(any(p["name"] == "Pin2Root" for p in self.app.pins_data))

    def test_create_route_from_selection_sufficient_pins(self):
        # Mock pins_data
        pin1_var = self.MockBooleanVar.return_value; pin1_var.get.return_value = True
        pin2_var = self.MockBooleanVar.return_value; pin2_var.get.return_value = True
        self.app.pins_data = [
            {"name": "PinA", "coords_original": (1,1,0), "coords_map": (1,1), "tk_var": pin1_var, "select_order": 1, "source": "s1"},
            {"name": "PinB", "coords_original": (2,2,0), "coords_map": (2,2), "tk_var": pin2_var, "select_order": 2, "source": "s1"},
        ]
        self.app.route_name_entry.get.return_value = "Test Route"
        self.app.route_color_combo.get.return_value = "rojo" # User-facing name
        
        self.app.create_route_from_selection()

        self.assertEqual(len(self.app.routes_data), 1)
        route = self.app.routes_data[0]
        self.assertEqual(route["name"], "Test Route")
        self.assertEqual(route["color"], "red") # Internal color name
        self.assertEqual(len(route["kml_coords"]), 2)
        self.app.map_widget.set_path.assert_called_once()
        MOCK_MODULES['tkinter.messagebox'].showinfo.assert_called_once()
        self.app._apply_theme.assert_called() # Check if theme is reapplied

    def test_create_route_from_selection_insufficient_pins(self):
        pin1_var = self.MockBooleanVar.return_value; pin1_var.get.return_value = True
        self.app.pins_data = [
            {"name": "PinA", "coords_original": (1,1,0), "coords_map": (1,1), "tk_var": pin1_var, "select_order": 1, "source": "s1"},
        ]
        self.app.create_route_from_selection()
        
        self.assertEqual(len(self.app.routes_data), 0)
        MOCK_MODULES['tkinter.messagebox'].showwarning.assert_called_once()
        # _apply_theme might still be called if the method reaches certain points,
        # depending on its placement in the original code. If it's in a finally or after messagebox,
        # this assertion might need adjustment or the original code might need guards.
        # For now, assuming it's not called if the main action isn't performed.

    def test_save_routes_to_kml_no_routes(self):
        self.app.routes_data = []
        self.app.save_routes_to_kml()
        
        MOCK_MODULES['tkinter.filedialog'].asksaveasfilename.assert_not_called()
        MOCK_MODULES['tkinter.messagebox'].showinfo.assert_called_with("Sin Rutas", "No hay rutas creadas para guardar.")

    def test_save_routes_to_kml_with_routes(self):
        self.app.routes_data = [
            {"name": "Route1", "kml_coords": [(1,1,0), (2,2,0)], "color": "red"}
        ]
        MOCK_MODULES['tkinter.filedialog'].asksaveasfilename.return_value = "dummy_path.kml"
        
        mock_kml_instance = MOCK_MODULES['simplekml'].Kml.return_value
        
        self.app.save_routes_to_kml()
        
        MOCK_MODULES['simplekml'].Kml.assert_called_once_with(name="Rutas Generadas")
        mock_kml_instance.newlinestring.assert_called_once_with(name="Route1", coords=[(1,1,0), (2,2,0)])
        mock_kml_instance.save.assert_called_once_with("dummy_path.kml")
        MOCK_MODULES['tkinter.messagebox'].showinfo.assert_called_with("Guardado Exitoso", "Rutas guardadas en 'dummy_path.kml' usando SimpleKML.")

    def test_select_all_deselect_all_pins(self):
        mock_vars = [MagicMock(spec=MOCK_MODULES['tkinter'].BooleanVar) for _ in range(3)]
        self.app.pins_data = [
            {"tk_var": mock_vars[0]},
            {"tk_var": mock_vars[1]},
            {"tk_var": mock_vars[2]},
        ]
        
        self.app.select_all_pins()
        for var_mock in mock_vars:
            var_mock.set.assert_called_with(True)
        
        self.app.deselect_all_pins()
        for var_mock in mock_vars:
            var_mock.set.assert_called_with(False)

    def test_create_routes_from_all(self):
        pin1_var = self.MockBooleanVar.return_value
        pin2_var = self.MockBooleanVar.return_value
        pin3_var = self.MockBooleanVar.return_value
        pin4_var = self.MockBooleanVar.return_value
        
        self.app.pins_data = [
            {"name": "P1S1", "coords_original": (1,1,0), "coords_map": (1,1), "tk_var": pin1_var, "source": "sourceA.kmz"},
            {"name": "P2S1", "coords_original": (2,1,0), "coords_map": (1,2), "tk_var": pin2_var, "source": "sourceA.kmz"},
            {"name": "P1S2", "coords_original": (3,3,0), "coords_map": (3,3), "tk_var": pin3_var, "source": "sourceB.kmz"},
            {"name": "P2S2", "coords_original": (4,3,0), "coords_map": (3,4), "tk_var": pin4_var, "source": "sourceB.kmz"},
            {"name": "P1S3_single", "coords_original": (5,5,0), "coords_map": (5,5), "tk_var": MagicMock(), "source": "sourceC_single.kmz"},
        ]
        
        self.app.create_routes_from_all()
        
        self.assertEqual(len(self.app.routes_data), 2) # sourceC has only 1 pin
        
        route_names = [r["name"] for r in self.app.routes_data]
        self.assertIn("Ruta sourceA.kmz", route_names)
        self.assertIn("Ruta sourceB.kmz", route_names)
        self.assertNotIn("Ruta sourceC_single.kmz", route_names)
        
        self.assertEqual(self.app.map_widget.set_path.call_count, 2)
        MOCK_MODULES['tkinter.messagebox'].showinfo.assert_called_with("Rutas Automáticas", "Se crearon 2 rutas automáticas.")


if __name__ == '__main__':
    unittest.main(argv=['first-arg-is-ignored'], exit=False)

# Ensure KML_NS is available for create_mock_element if run directly for some reason
# KML_NS_TEST = "{http://www.opengis.net/kml/2.2}" if 'KML_NS_TEST' not in globals() else KML_NS_TEST
