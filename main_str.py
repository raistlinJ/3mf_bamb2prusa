###
# main_str.py
# This script provides a GUI for converting Bambu 3mf files to Prusa 3mf files.
# It allows users to select input and output files, decompress the input zip file,
# process the 3mf files, and generate a new 3mf file with the converted content.
# The script uses lxml for XML parsing and tkinter for the GUI.
# @Author: Jaime C. Acosta
# @Date: 2025-08-09
# @Version: 1.0
# @License: GPL 3.0
# @Description: A GUI application to convert Bambu 3mf files to Prusa 3mf files.
###
import os
import shutil
import tempfile
import zipfile
import re
import io
import logging
from tkinter import Tk, Label, Button, filedialog
import lxml.etree as ET
from pathlib import Path

class ZipProcessorGUI:

    def __init__(self, master):
        logging.debug("Initializing ZipProcessorGUI")
        self.master = master
        master.title("Bambu2Prusa 3mf Processor")

        self.label = Label(master, text="Select input and output 3mf files.")
        self.label.pack(pady=10)

        self.select_input_button = Button(master, text="Select Input Bambu 3mf", command=self.select_input)
        self.select_input_button.pack(pady=5)

        self.select_output_button = Button(master, text="Select Output Prusa 3mf", command=self.select_output)
        self.select_output_button.pack(pady=5)

        self.process_button = Button(master, text="Process", command=self.bambu3mf2prusa3mf)
        self.process_button.pack(pady=10)

        self.status_label = Label(master, text="")
        self.status_label.pack(pady=5)

        self.input_file = ""
        self.output_file = ""

        # Define paths for templates and directories
        self.template_paths = {}
        self.template_paths['models_template'] = "3mf_template/3D/3dmodel_template.xml"
        self.template_paths['3D'] = "3mf_template/3D/"
        self.template_paths['.rels_template'] = "3mf_template/_rels/.rels_template.xml"
        self.template_paths['.rels'] = "3mf_template/_rels/"
        self.template_paths['Content_Types_template'] = "3mf_template/[Content_Types].xml"
        self.template_paths['Metadata'] = "3mf_template/Metadata/"

        self.temp_3mf_dir = tempfile.TemporaryDirectory().name

        self.bambu_model_paths = []
        #contains output object file names and the object ids within those files
        self.prusa_model_paths = {}

    def select_input(self):
        logging.debug("Selecting input file")
        # Use filedialog to select a 3mf file
        self.input_file = filedialog.askopenfilename(filetypes=[("3mf files", "*.3mf")])
        self.status_label.config(text=f"Input file selected: {os.path.basename(self.input_file)}")

    def select_output(self):
        logging.debug("Selecting output file")
        # Use filedialog to select an output file
        self.output_file = filedialog.asksaveasfilename(defaultextension=".3mf", filetypes=[("3mf files", "*.3mf")])
        self.status_label.config(text=f"Output file selected: {os.path.basename(self.output_file)}")

    def decompress_zip(self, input_file=None):
        # Decompress the zip file to a temporary directory
        logging.debug("Decompressing zip file")
        if input_file == None:
            input_file = self.input_file
        # Check if input file is provided
        if not input_file:
            self.status_label.config(text="Please provide both input and output files.")
            return
        # Create a temporary directory for extraction
        tempdir = tempfile.TemporaryDirectory().name

        # Unzip the input file
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(tempdir)
        # return the temporary directory path that contains the extracted files
        return tempdir
            
    def bambu3mf2prusa3mf(self, input_file=None, output_file=None, extracted_path=None):
        logging.debug("Converting Bambu 3mf to Prusa 3mf")
        # Check if input and output files are provided
        if input_file == None:
            input_file = self.input_file
        if output_file == None:
            output_file = self.output_file

        if not input_file or not output_file:
            self.status_label.config(text="Please provide both input and output files.")
            return
        #if we haven't specified an extracted path, we will decompress the zip file
        if extracted_path==None:
            extracted_path = self.decompress_zip(input_file)
        objects_path = os.path.join(extracted_path,"3D","Objects")
        # Check if the objects directory exists
        if os.path.exists(objects_path):
            # Parse all .model files
            self.bambu_model_paths = list(Path(objects_path).rglob("*.model"))
            if self.bambu_model_paths == None or self.bambu_model_paths == None:
                logging.error("No model files found")
                return 
            
        # Convert each model file to Prusa format
        obj_IDs_Elements = []
        for bmodel_path in self.bambu_model_paths:
            filename, obj_IDs_Element = self.model_convert_re(bmodel_path)
            obj_IDs_Elements.append([filename, obj_IDs_Element])

        #inject the objects into the Prusa model template
        prusamodel_filenames = []
        for filename, obj_IDs_Element in obj_IDs_Elements:
            final_prusamodel = self.inject_bobject2pobject(obj_IDs_Element)
            #final_prusamodels.append([filename, final_prusamodel])
            #logging.debug(ET.tostring(final_prusamodel, encoding='utf-8', pretty_print=True))
            self.write_prusa_model(filename, final_prusamodel)
            prusamodel_filenames.append(filename)

        # Write the final Prusa model files
        self.generate3mf_file(prusamodel_filenames, output_file)

    def model_convert_re(self, bmodel_path):
        logging.debug(f"Processing model file: {bmodel_path}")
        # Check if the file exists
        if not os.path.exists(bmodel_path):
            logging.error(f"File not found: {bmodel_path}")
            return None, None
        
        relevant_objects = {}
        # convert the bambu model file to a prusa model file using regex and xml parsing
        try:
            objects = {}
            #we don't know what encoding is used for the xml string, so we force it to utf-8
            logging.debug(f"Reading model file: {bmodel_path}")
            with open(bmodel_path, encoding="utf-8") as f:
                content = f.read()
            #Add any necessary namespaces and remove any unwanted attributes; if file formats change, this may need to be updated
            #remove any namespaces since we're doing direct string replacements
            rem_xmlns = re.sub(r'xmlns=[^=]+"', '', content)
            #remove any p:UUID attributes
            rem_pUUID = re.sub(r'p:UUID[^"]+"[^"]+"', '', rem_xmlns)
            #remove any encoding attributes, since lxml only likes utf-8
            rem_encoding = re.sub("encoding=\"[0-9A-Z\\-]*\"", "", rem_pUUID)
            #replace the paint attribute with the slic3rpe namespace
            rem_paint_color = rem_encoding.replace("paint_color", "slic3rpe:mmu_segmentation")
            #replace the model tag with a new one that has the correct namespaces for prusa format
            add_slic3rpe = re.sub(r'<model[ ].*">', r'<model unit="millimeter" xml:lang="en-US" xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02" xmlns:slic3rpe="http://schemas.slic3r.org/3mf/2017/06">', rem_paint_color)
            #remove the paint_seam attribute, since it is not allowed in Prusa format
            rem_paint_seam = re.sub("paint_seam=\"[0-9A-Z]*\"", "", add_slic3rpe)
            logging.debug("Parsing XML content")
            #parse the xml content and find the objects
            bambu_tree = ET.fromstring(rem_paint_seam)
            objects = bambu_tree.findall(".//{*}resources/{*}object")

            #for each object, check if it is of type "model" and add it to the relevant_objects dictionary; these are the only object types that are allowed in prousa format
            for object in objects:
                if object.attrib['type'] == "model":
                    relevant_objects[object.attrib['id']] = object
                logging.debug(f"Object type {object.attrib['type']} | id {object.attrib['id']}: ")
            
        except FileNotFoundError:
            logging.error(f"Error: File '{bmodel_path}' not found.")
            return
        except Exception as e:
            logging.error(f"An error occurred: {e}")

        # take only the basename of the bmodel_path to use as the filename in the prusa model
        model_filename = os.path.basename(bmodel_path)

        return model_filename, relevant_objects
        
    def inject_bobject2pobject(self, bobjects):
        logging.debug("Injecting Bambu objects into Prusa model template")
        if not bobjects:
            logging.warning("No objects to inject into the template. Will use empty template.")
        # Open the Prusa model template file
        with open(self.template_paths['models_template'], 'r', encoding='utf-8') as f:
            template_content = f.read()

        #inject object into template file
        try:
            logging.debug("Injecting objects into the Prusa model template")
            #create build element so that we can add items under it for each object
            #build_element = ET.fromstring("<build></build>")
            #read the template file
            tree = ET.parse(self.template_paths['models_template'])
            model = tree.getroot()
            build = model.find(".//{*}build")
            for bobject in bobjects:
                build.append(ET.Element("item", objectid=bobject, transform="0.799151571 0 0 0 0.799151571 0 0 0 0.799151571 184.67373 221.31425 1.61151839", printable="1"))

            # Convert the tree back to a string so we can modify using regex (much faster than using lxml for this)
            template_content = ET.tostring(tree, encoding='unicode')
            logging.debug("Model root element found")
            for bobject in bobjects:
                logging.debug(f"Adding object {bobject} to template content")
                template_content = re.sub("<resources>", "<resources>\n"+str(ET.tostring(bobjects[bobject], encoding='unicode', pretty_print=True)), template_content)

            if template_content is None:
                logging.error("Template content is None, cannot proceed with injection.")
                return None
            return ET.fromstring(template_content)

        except FileNotFoundError:
            logging.error(f"Error: File '{self.template_paths['models_template']}' not found.")
            self.status_label.config(text=f"Error reading {self.template_paths['models_template']}: {e}")
            return
        except Exception as e:
            logging.error(f"An error occurred: {e}")

    def compress_zip(self, ifolder_path, output_file=None):
        logging.debug("Compressing files into zip")
        # Check if the input folder path is provided
        if not ifolder_path:
            self.status_label.config(text="No input folder provided for compression.")
            return
        if output_file is None:
            output_file = self.output_file
        # Re-zip contents into the output file
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for foldername, subfolders, filenames in os.walk(ifolder_path):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, ifolder_path)
                    zip_out.write(file_path, arcname)
        logging.info(f"Compressed files into {output_file}")
        self.status_label.config(text=f"Output file created: {os.path.basename(output_file)}")

    def write_prusa_model(self, filename, prusa_model):
        logging.debug("Writing Prusa object")
        # This function is a placeholder for any additional processing needed
        tempdir = self.temp_3mf_dir
        objects_dir = os.path.join(tempdir, "3D", "Objects")
        os.makedirs(objects_dir, exist_ok=True)

        try:
            ###--3D/Objects/3dmodel.xml---###
            # Create the 3D Objects directory and copy the model files
            model_path = os.path.join(objects_dir, filename)
            if prusa_model == None or prusa_model == []:
                logging.warning("Prusa model is empty, writing empty object.")
                #return
            prusa_model.getroottree().write(model_path, encoding='utf-8', xml_declaration=True, pretty_print=True)
        except Exception as e:
            logging.error(f"An error occurred while writing Prusa object: {e}")

    def generate3mf_file(self, final_prusamodels, output_file=None):
        logging.debug("Generating 3mf file structure")
        # Check if the output file is provided
        if output_file is None:
            output_file = self.output_file
        if not output_file:
            self.status_label.config(text="Please provide an output file.")
            return
        # Check if there are any models to generate the 3mf file structure
        if not final_prusamodels:
            logging.error("No models to generate 3mf file structure.")
            self.status_label.config(text="No models to generate 3mf file structure.")
            return
        try:
            # Create the temporary directory for the 3mf file structure
            self.status_label.config(text="Generating 3mf file structure...")
            tempdir = self.temp_3mf_dir
            # Create the necessary directories
            rels_dir = os.path.join(tempdir, "_rels")
            os.makedirs(rels_dir, exist_ok=True)
            metadata_dir = os.path.join(tempdir, "Metadata")
            os.makedirs(metadata_dir, exist_ok=True)


            ###--[Content-Types].xml---###
            # Copy the template files to the output directory
            shutil.copy(self.template_paths['Content_Types_template'], os.path.join(tempdir, "[Content_Types].xml"))


            ###--_rels/.rels---###
            # Create the relationships file
            rels_ET = ET.parse(self.template_paths['.rels_template'])
            rels_tree = rels_ET.getroot()
            relationship_number = 1
            for model in final_prusamodels:
                # Add a relationship for the model
                rel = ET.fromstring(f'<Relationship Target="/3D/Objects/{model}" Id="rel-{relationship_number}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/3dmodel"/>')
                relationship_number += 1
                rels_tree.append(rel)
            # Write the relationships file
            rels_path = os.path.join(rels_dir, ".rels")
            rels_ET.write(rels_path, encoding='utf-8', xml_declaration=True, pretty_print=True)

            # Add the Metadata files if they exist
            if os.path.exists(self.template_paths['Metadata']):
                for file in os.listdir(self.template_paths['Metadata']):
                    shutil.copy(os.path.join(self.template_paths['Metadata'], file), os.path.join(tempdir, "Metadata", file))

            self.compress_zip(tempdir, output_file)
        except Exception as e:
            logging.error(f"An error occurred while generating 3mf file structure: {e}")

def main():
    root = Tk()
    app = ZipProcessorGUI(root)
    root.mainloop()
    # If the GUI is not needed, you can run the conversion directly
    # b3mf = f"input.3mf"
    # output = f"output.3mf"
    # a = ZipProcessorGUI(root)
    # a.bambu3mf2prusa3mf(b3mf, output)

if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    main()
