"""
This module provides a robust class to interact with Aerotech Automation1 DLLs
needed for executing a Machine Setup calculation.
This class will generate a new MCD file with calculated parameters based on
either stage configuration options, or an existing MCD file. The resulting MCD
will have parameters calculated as if it were processed through Machine Setup. 
It uses an explicit initialize() method to ensure .NET assemblies are loaded
at the correct time, preventing initialization errors.
"""
from pythonnet import load
load("coreclr")
import os
import sys
import json
from tkinter import filedialog, ttk, messagebox
import xml.etree.ElementTree as ET

# Import System for Type.GetType
import System
from System.Collections.Generic import List
from System import String

import clr

sys.dont_write_bytecode = True

class AerotechController:
    """
    A class to encapsulate the functionality of Aerotech's machine 
    configuration DLLs. It separates object creation from the sensitive
    .NET assembly loading and provides a clean API for core functions.

    Usage:
       
        # --- Workflow 1: JSON Specs -> Calculated MCD Object ---
        mcd_obj, _ = controller.convert_to_mcd(specs_dict, ...)
        calculated_mcd, _ = controller.calculate_parameters(mcd_obj)

        # --- Workflow 2: MCD File -> JSON File ---
        controller.convert_to_json("path/to/input.mcd", "path/to/output.json")
    """

    def __init__(self, mcd_name=None):
        """
        Initialize the AerotechController instance.

        Sets up paths for required DLLs, configuration files, and templates.
        Dynamically locates the latest installed Automation1 version and validates
        the presence of required dependencies. Does NOT load any .NET assemblies.

        Args:
            mcd_name (str, optional): Optional name for the MCD file. If not provided,
                the stage type will be used for naming output files.

        Raises:
            FileNotFoundError: If required configuration or DLL directories are missing.
            Displays a messagebox warning if Automation1 2.11 or newer is not found.
        """
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Automatically determine config_manager_path based on base_dir
        config_manager_path = os.path.join(self.base_dir, "GenerateMCD Assets", "System.Configuration.ConfigurationManager.8.0.0", "lib", "netstandard2.0")
        if not os.path.exists(config_manager_path):
            raise FileNotFoundError(f"ConfigurationManager path not found: {config_manager_path}")
        
        # Dynamically find the latest Automation1 version folder
        automation1_root = r"C:\Program Files (x86)\Aerotech\Controller Version Selector\Bin\Automation1"
        latest_version = None
        if os.path.exists(automation1_root):
            # List all subfolders that look like version numbers (e.g., "2.11.0")
            version_folders = [
                name for name in os.listdir(automation1_root)
                if os.path.isdir(os.path.join(automation1_root, name)) and name[0].isdigit()
            ]
            if version_folders:
                # Sort by version (major.minor.patch), descending
                # Use packaging.version if available (preferred for modern Python environments)
                try:
                    from packaging.version import Version
                    version_folders.sort(key=Version, reverse=True)
                except ImportError:
                    # Fallback: use tuple conversion for basic version sorting
                    def version_tuple(v):
                        return tuple(int(x) for x in v.split('.') if x.isdigit())
                    version_folders.sort(key=version_tuple, reverse=True)
                latest_version = version_folders[0]
        if not latest_version:
            message = (
                "This class only works with Automation1 2.11 or newer.\n"
                "Please install Automation1 2.11.x or later."
            )
            try:
                messagebox.showwarning("Automation1 Version Not Found", message)
            except Exception:
                print("Warning: " + message)
            aerotech_dll_path = None
        else:
            aerotech_dll_path = os.path.join(
                automation1_root, latest_version, "release", "Bin"
            )
            if not os.path.exists(aerotech_dll_path):
                raise FileNotFoundError(f"Aerotech DLL path not found: {aerotech_dll_path}")
        
        self.working_dir = os.getcwd()
        self.aerotech_dll_path = aerotech_dll_path
        self.config_manager_path = config_manager_path

        self.mcd_name = mcd_name

        # Define paths using the provided base_dir
        self.template_path = os.path.join(self.base_dir, "GenerateMCD Assets", "MS_Template.json")
        self.working_json_path = os.path.join(self.base_dir, "GenerateMCD Assets", "WorkingTemplate.json")

        self.McdFormatConverter = None
        self.MachineControllerDefinition = None
        self.JObject = None
        self.initialized = False

    def initialize(self):
        """
        Load all required .NET assemblies and types for Automation1 interaction.

        This method must be called before any other method that interacts with the DLLs.
        It sets up the .NET runtime environment, loads all necessary assemblies, and
        initializes type handles for later use.

        Raises:
            RuntimeError: If any required .NET assembly or type cannot be loaded.
        """
        if self.initialized:
            print("Controller already initialized.")
            return

        os.environ["PATH"] = self.aerotech_dll_path + ";" + os.environ["PATH"]
        os.add_dll_directory(self.aerotech_dll_path)
        
        try:
            clr.AddReference(os.path.join(self.aerotech_dll_path, "Newtonsoft.Json.dll"))
            clr.AddReference(os.path.join(self.config_manager_path, "System.Configuration.ConfigurationManager.dll"))
            clr.AddReference(os.path.join(self.aerotech_dll_path, "Aerotech.Automation1.Applications.Core.dll"))
            clr.AddReference(os.path.join(self.aerotech_dll_path, "Aerotech.Automation1.Applications.Interfaces.dll"))
            clr.AddReference(os.path.join(self.aerotech_dll_path, "Aerotech.Automation1.Applications.Shared.dll"))
            clr.AddReference(os.path.join(self.aerotech_dll_path, "Aerotech.Automation1.DotNetInternal.dll"))
            clr.AddReference(os.path.join(self.aerotech_dll_path, "Aerotech.Automation1.Applications.Wpf.dll"))

            import Newtonsoft.Json.Linq
            self.JObject = Newtonsoft.Json.Linq.JObject

            type_name1 = "Aerotech.Automation1.Applications.Wpf.McdFormatConverter, Aerotech.Automation1.Applications.Wpf"
            type_name2 = "Aerotech.Automation1.DotNetInternal.MachineControllerDefinition, Aerotech.Automation1.DotNetInternal"
            
            self.McdFormatConverter = System.Type.GetType(type_name1)
            self.MachineControllerDefinition = System.Type.GetType(type_name2)

            if self.McdFormatConverter is None or self.MachineControllerDefinition is None:
                raise TypeError("Could not load required .NET types. Check DLL versions and names.")
            
            self.initialized = True

        except Exception as e:
            self.initialized = False
            raise RuntimeError(f"Failed to initialize controller: {e}")

    def _check_initialized(self):
        """
        Ensure the controller has been initialized.

        Raises:
            RuntimeError: If the controller has not been initialized via initialize().
        """
        if not self.initialized:
            raise RuntimeError("Controller has not been initialized. Please call controller.initialize() first.")

    def _update_json_config(self, specs_dict, stage_type=None, axis=None):
        """
        Update the working JSON configuration file with new specifications.

        Modifies a template JSON file with the provided stage configuration options,
        stage type, and axis name. Writes the updated configuration to a working file.

        Args:
            specs_dict (dict): Dictionary of configuration options to update.
            stage_type (str, optional): Name of the stage type to set.
            axis (str, optional): Name of the axis to set.

        Raises:
            KeyError: If required keys are missing in the template JSON.
        """
        with open(self.template_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        mech_products = data.get("MechanicalProducts")
        if not mech_products:
            raise KeyError("MechanicalProducts not found in JSON.")
        
        mech_product = mech_products[0]
        mech_product.setdefault("ConfiguredOptions", {}).update(specs_dict)

        if stage_type:
            mech_product["Name"] = stage_type
            mech_product["DisplayName"] = stage_type

        interconnected_axes = data.get("InterconnectedAxes")
        if interconnected_axes:
            inter_axis = interconnected_axes[0]
            if axis:
                inter_axis["Name"] = axis
            if stage_type and "MechanicalAxis" in inter_axis:
                inter_axis["MechanicalAxis"]["DisplayName"] = stage_type

        with open(self.working_json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        print(f"Configuration updated in {self.working_json_path}")

    def _read_mcd_from_file(self, mcd_path):
        """
        Read an MCD file from disk and return the corresponding .NET object.

        Also checks that the Automation1 software version is 2.11 or newer.

        Args:
            mcd_path (str): Path to the MCD file.

        Returns:
            .NET MachineControllerDefinition object.

        Raises:
            FileNotFoundError: If the MCD file does not exist.
            RuntimeError: If the Automation1 version is not 2.11 or newer.
        """
        self._check_initialized()
        if not os.path.exists(mcd_path):
            raise FileNotFoundError(f"MCD file not found at: {mcd_path}")
            
        read_from_file = self.MachineControllerDefinition.GetMethod("ReadFromFile")
        mcd = read_from_file.Invoke(None, [mcd_path])

        version = mcd.SoftwareVersion

        # Check if version is at least 2.11
        def is_version_supported(ver_str):
            try:
                # Only compare major.minor
                parts = ver_str.split('.')
                major = int(parts[0])
                minor = int(parts[1]) if len(parts) > 1 else 0
                return (major > 2) or (major == 2 and minor >= 11)
            except Exception:
                return False

        if not is_version_supported(str(version)):
            message = (
                f"This class is only supported for Automation1 2.11 or newer.\n"
                f"Detected version: {version}\n"
            )
            try:
                messagebox.showwarning("Unsupported Automation1 Version", message)
            except Exception:
                print("Warning: " + message)
            # Exit the method/class by raising an exception
            raise RuntimeError("Unsupported Automation1 version: " + str(version))
        return mcd

    def convert_to_mcd(self, specs_dict=None, stage_type=None, axis=None, workflow=None):
        """
        Create a working JSON config from a template and convert it to an MCD object.

        Updates the working JSON configuration with the provided specs, then converts
        it to a .NET MachineControllerDefinition object using the Automation1 DLLs.

        Args:
            specs_dict (dict, optional): Configuration options to update in the template.
            stage_type (str, optional): Name of the stage type.
            axis (str, optional): Name of the axis.
            workflow (str, optional): If 'wf2', saves an uncalculated MCD file.

        Returns:
            tuple: (mcd_obj, mcd_path, warnings)
                mcd_obj: The .NET MCD object.
                mcd_path: Path to the generated MCD file.
                warnings: List of warning strings from the conversion process.
        """
        self._check_initialized()
            
        if specs_dict:
            self._update_json_config(specs_dict, stage_type, axis)
        
            with open(self.working_json_path, "r", encoding="utf-8") as f:
                json_str = f.read()
        else:
            full_mcd_json = os.path.join(self.working_dir, f"{stage_type}.json")
            with open(full_mcd_json, "r", encoding="utf-8") as f:
                json_str = f.read()
        jobject = self.JObject.Parse(json_str)
        warnings = List[String]()

        convert_method = self.McdFormatConverter.GetMethod("ConvertToMcd")
        mcd_obj = convert_method.Invoke(None, [jobject, warnings])
        
        #if workflow == 'wf2':
        if self.mcd_name is None:
            mcd_path = os.path.join(self.working_dir, f"Uncalculated_{stage_type}.mcd")
        else:
            mcd_path = os.path.join(self.working_dir, f"Uncalculated_{self.mcd_name}.mcd")
        mcd_obj.WriteToFile(mcd_path)

        # Clean up temporary working template
        #self._cleanup_working_template()

        return mcd_obj, mcd_path, list(warnings)

    def _cleanup_working_template(self):
        """
        Delete the temporary working JSON template file if it exists.

        Prints a warning if the file cannot be deleted.
        """
        try:
            if os.path.exists(self.working_json_path):
                os.remove(self.working_json_path)

        except Exception as e:
            print(f"Warning: Could not clean up WorkingTemplate.json: {e}")

    def convert_to_json(self, mcd_path, output_json_path):
        """
        Convert an MCD file to a JSON file.

        Reads an MCD file from disk, converts it to a JSON representation using
        the Automation1 DLLs, and writes the result to the specified output path.

        Args:
            mcd_path (str): Path to the input MCD file.
            output_json_path (str): Path to write the output JSON file.

        Returns:
            list: List of warning strings from the conversion process.

        Raises:
            FileNotFoundError: If the MCD file does not exist.
        """
        self._check_initialized()
        mcd_obj = self._read_mcd_from_file(mcd_path)
        
        warnings = List[String]()
        convert_method = self.McdFormatConverter.GetMethod("ConvertToJson")
        json_obj = convert_method.Invoke(None, [mcd_obj, warnings])

        with open(output_json_path, 'w', encoding='utf-8') as f:
            f.write(json_obj.ToString())

        return list(warnings)

    def calculate_parameters(self, specs_dict=None, stage_type=None, axis=None):
        """
        Create an MCD object from specs and calculate its parameters.

        This is the primary method for the JSON-to-calculated-MCD workflow.
        Converts the provided specs to an MCD object, then calculates and saves
        the parameters using the Automation1 DLLs.

        Args:
            specs_dict (dict, optional): Configuration options to update in the template.
            stage_type (str, optional): Name of the stage type.
            axis (str, optional): Name of the axis.

        Returns:
            tuple: (calculated_mcd, all_warnings, mcd_path)
                calculated_mcd: The calculated .NET MCD object.
                all_warnings: List of warning strings from both conversion and calculation.
                mcd_path: Path to the generated MCD file.
        """
        if self.mcd_name is None:
            mcd_path = os.path.join(self.working_dir, f"Calculated_{stage_type}.mcd")

        self._check_initialized()
        
        # Step 1: Convert specs to an initial MCD object.

        mcd_obj, _, conversion_warnings = self.convert_to_mcd(specs_dict=specs_dict, stage_type=stage_type, axis=axis)
        if conversion_warnings:
            print("Warnings during conversion:", conversion_warnings)
        
        # Step 2: Calculate parameters on the new MCD object.
        warnings = List[String]()
        calculate_method = self.McdFormatConverter.GetMethod("CalculateParameters")
        calculated_mcd = calculate_method.Invoke(None, [mcd_obj, warnings])
        calculation_warnings = list(warnings)
        if calculation_warnings:
            print("Warnings during calculation:", calculation_warnings)
        
        calculated_mcd.WriteToFile(mcd_path)

        # Combine warnings from both steps.
        all_warnings = conversion_warnings + calculation_warnings

        return calculated_mcd, all_warnings, mcd_path

    def calculate_from_current_mcd(self, mcd_path):
        """
        Calculate parameters for an existing MCD object.

        Args:
            mcd_path: Filepath to the MCD used to calculate new parameters.

        Returns:
            tuple: (calculated_mcd, mcd_path, all_warnings)
                calculated_mcd: The recalculated .NET MCD object.
                mcd_path: Path to the recalculated MCD file.
                all_warnings: List of warning strings from the calculation.
        """
        # Creating MCD object
        mcd_obj = self._read_mcd_from_file(mcd_path)

        warnings = List[String]()
        calculate_method = self.McdFormatConverter.GetMethod("CalculateParameters")
        calculated_mcd = calculate_method.Invoke(None, [mcd_obj, warnings])
        calculation_warnings = list(warnings)
        if calculation_warnings:
            print("Warnings during calculation:", calculation_warnings)
        
        if self.mcd_name is None:
            mcd_path = os.path.join(self.working_dir, f"Recalculated.mcd")
            calculated_mcd.WriteToFile(mcd_path)

        # Combine warnings from both steps.
        all_warnings = calculation_warnings
        
        return calculated_mcd, mcd_path, all_warnings

    def inspect_mcd_object(self, calculated_mcd_object):
        """
        Inspect a calculated MCD object and extract parameter information.

        Examines the ConfigurationFiles property of the MCD object, prints
        information about the 'Parameters' entry, and attempts to extract
        servo loop and feedforward parameters from the XML content.

        Args:
            calculated_mcd_object: The .NET MCD object to inspect.

        Returns:
            tuple: (servo_params, feedforward_params)
                servo_params: Dictionary of servo loop parameters by axis.
                feedforward_params: Dictionary of feedforward parameters by axis.
            or None if parameters cannot be extracted.
        """
        dotnet_type = calculated_mcd_object.GetType()

        # Examine ConfigurationFiles
        config_files_prop = dotnet_type.GetProperty("ConfigurationFiles")
        config_files = config_files_prop.GetValue(calculated_mcd_object, None)

        parameters_filedata = None
        if config_files is not None:
            for item in config_files:
                key = getattr(item, "Key", None)
                value = getattr(item, "Value", None)
                key_str = str(key) if key is not None else None

                if key_str == "Parameters":
                    parameters_filedata = value
            if parameters_filedata is not None:

                for prop in parameters_filedata.GetType().GetProperties():
                    val = prop.GetValue(parameters_filedata, None)
                
                content_prop = parameters_filedata.GetType().GetProperty("Content")
                content_bytes = content_prop.GetValue(parameters_filedata, None)

                # Convert .NET byte[] to Python bytes
                if content_bytes is not None:
                    # pythonnet exposes byte[] as a sequence of ints
                    py_bytes = bytes(bytearray(content_bytes))

                    # Try to decode as UTF-8 text (if it's text)
                    try:
                        text = py_bytes.decode('utf-8')
                        servo_params = self.extract_servo_loop_parameters_from_xml(text)
                        feedforward_params = self.extract_feedforward_parameters_from_xml(text)
                        
                        return servo_params, feedforward_params
                    except Exception as e:
                        print(f"\nCould not decode Parameters Content as UTF-8 text: {e}")
                else:
                    print("Parameters Content is None.")
            else:
                print("Parameters entry not found in ConfigurationFiles.")
        else:
            print("ConfigurationFiles is None or empty.")
    
    def extract_servo_loop_parameters_from_xml(self, xml_text):
        """
        Parse XML and extract all ServoLoop parameters for all axes.

        Args:
            xml_text (str): XML string containing axis parameters.

        Returns:
            dict: {axis_index: [{name, value}, ...], ...}
        """
        axis_servo_params = {}
        root = ET.fromstring(xml_text)
        for axis_elem in root.findall('.//Axes/Axis'):
            axis_index = axis_elem.attrib.get('Index')
            params = []
            for p in axis_elem.findall('P'):
                param_name = p.attrib.get('n', '')
                if param_name.startswith('ServoLoop'):
                    param = {
                        'name': param_name,
                        'value': p.text
                    }
                    params.append(param)
            if params:
                axis_servo_params[axis_index] = params
        return axis_servo_params

    def extract_feedforward_parameters_from_xml(self, xml_text):
        """
        Parse XML and extract all Feedforward parameters for all axes.

        Args:
            xml_text (str): XML string containing axis parameters.

        Returns:
            dict: {axis_index: [{name, value}, ...], ...}
        """
        axis_feedforward_params = {}
        root = ET.fromstring(xml_text)
        for axis_elem in root.findall('.//Axes/Axis'):
            axis_index = axis_elem.attrib.get('Index')
            params = []
            for p in axis_elem.findall('P'):
                param_name = p.attrib.get('n', '')
                if param_name.startswith('Feedforward'):
                    param = {
                        'name': param_name,
                        'value': p.text
                    }
                    params.append(param)
            if params:
                axis_feedforward_params[axis_index] = params
        return axis_feedforward_params


if __name__ == "__main__":
    # This is a module, not meant to be run directly
    # Import and use the AerotechController class in other scripts
    print("GenerateMCD module loaded. Use AerotechController class in your scripts.")


