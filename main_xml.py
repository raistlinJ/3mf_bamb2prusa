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

        self.template_paths = {}
        self.template_paths['models_template'] = "3mf_template/3D/3dmodel_template.xml"
        self.template_paths['3D'] = "3mf_template/3D/"
        self.template_paths['.rels_template'] = "3mf_template/_rels/.rels_template.xml"
        self.template_paths['.rels'] = "3mf_template/_rels/"
        self.template_paths['Content_Types_template'] = "3mf_template/[Content_Types].xml"
        self.template_paths['Metadata'] = "3mf_template/Metadata/"

        self.temp_3mf_dir = "temp_3mf"

        self.bambu_model_paths = []
        #contains output object file names and the object ids within those files
        self.prusa_model_paths = {}

    def select_input(self):
        logging.debug("Selecting input file")
        self.input_file = filedialog.askopenfilename(filetypes=[("3mf files", "*.3mf")])
        self.status_label.config(text=f"Input file selected: {os.path.basename(self.input_file)}")

    def select_output(self):
        logging.debug("Selecting output file")
        self.output_file = filedialog.asksaveasfilename(defaultextension=".3mf", filetypes=[("3mf files", "*.3mf")])
        self.status_label.config(text=f"Output file selected: {os.path.basename(self.output_file)}")

    def decompress_zip(self, input_file=None):
        logging.debug("Decompressing zip file")
        if input_file == None:
            input_file = self.input_file

        if not input_file:
            self.status_label.config(text="Please provide both input and output files.")
            return

        tempdir = tempfile.TemporaryDirectory().name

        # Unzip the input file
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(tempdir)
        
        return tempdir
            
    def bambu3mf2prusa3mf(self, input_file=None, output_file=None, extracted_path=None):
        logging.debug("Converting Bambu 3mf to Prusa 3mf")
        if input_file == None:
            input_file = self.input_file
        if output_file == None:
            output_file = self.output_file

        if not input_file or not output_file:
            self.status_label.config(text="Please provide both input and output files.")
            return

        if extracted_path==None:
            extracted_path = self.decompress_zip(input_file)
        objects_path = os.path.join(extracted_path,"3D","Objects")
        if os.path.exists(objects_path):
            # Parse all .model files
            self.bambu_model_paths = list(Path(objects_path).rglob("*.model"))

            if self.bambu_model_paths == None or self.bambu_model_paths == None:
                logging.error("No model files found")
                return 
            
        obj_IDs_Elements = []
        for bmodel_path in self.bambu_model_paths:
            filename, obj_IDs_Element = self.model_convert_re(bmodel_path)
            obj_IDs_Elements.append([filename, obj_IDs_Element])

        prusamodel_filenames = []
        for filename, obj_IDs_Element in obj_IDs_Elements:
            final_prusamodel = self.inject_bobject2pobject(obj_IDs_Element)
            #final_prusamodels.append([filename, final_prusamodel])
            #logging.debug(ET.tostring(final_prusamodel, pretty_print=True))
            self.write_prusa_model(filename, final_prusamodel)
            prusamodel_filenames.append(filename)

        self.generate3mf_file(prusamodel_filenames)

    def model_convert_re(self, bmodel_path):
        logging.debug(f"Processing model file: {bmodel_path}")
        if not os.path.exists(bmodel_path):
            logging.error(f"File not found: {bmodel_path}")
            return None, None
        try:
            objects = {}
            #we don't know what encoding is used for the xml string, so we force it to utf-8
            with open(bmodel_path, encoding="utf-8") as f:
                content = f.read()
            temp = ET.parse(bmodel_path)
            rem_encoding = re.sub("encoding=\"[0-9A-Z\\-]*\"", "", content)
            rem_paint_color = rem_encoding.replace("paint_color", "slic3rpe:mmu_segmentation")
            rem_paint_seam = re.sub("paint_seam=\"[0-9A-Z]*\"", "", rem_paint_color)

            bambu_tree = ET.fromstring(rem_paint_seam)
            objects = bambu_tree.findall(".//{*}resources/{*}object")
            relevant_objects = {}
            for object in objects:
                if object.attrib['type'] == "model":
                    relevant_objects[object.attrib['id']] = object
                logging.debug(f"Object type {object.attrib['type']} | id {object.attrib['id']}: ")
            #filter out those that are of type "model"
        except FileNotFoundError:
            logging.error(f"Error: File '{bmodel_path}' not found.")
            return
        except Exception as e:
            logging.error(f"An error occurred: {e}")
        
        model_filename = os.path.basename(bmodel_path)

        return model_filename, relevant_objects
        
    def inject_bobject2pobject(self, bobjects):
        logging.debug("Injecting Bambu objects into Prusa model template")
        if not bobjects:
            logging.error("No objects to inject into the template.")
            return None
        #inject object into template file
        try:
            logging.debug("Injecting objects into the Prusa model template")
            #create build element so that we can add items under it for each object
            build_element = ET.fromstring("<build></build>")
            #read the template file
            tree = ET.parse(self.template_paths['models_template'])
            model = tree.getroot()
            model.append(build_element)
            logging.debug("Model root element found")
            for bobject in bobjects:
                logging.debug(f"Injecting object {bobject} into the model")
                #create a new object element for each object
                #get the resources tag
                logging.debug(f"Finding resources in template")
                resources = model.find(".//{*}resources")
                #add each of the objects as sub elements
                logging.debug(f"Adding object {bobject} to resources")
                resources.append(bobjects[bobject])


                #in the model tag, add a line for each object
                logging.debug(f"Adding item element for object {bobject}")
                #create an item element with the object id and transform

                item_element = ET.fromstring(f"<item objectid=\"{bobject}\" transform=\"0.799151571 0 0 0 0.799151571 0 0 0 0.799151571 184.67373 221.31425 1.61151839\" printable=\"1\"/>")
                build_element.append(item_element)
            return tree

        except FileNotFoundError:
            logging.error(f"Error: File '{self.template_paths['models_template']}' not found.")
            self.status_label.config(text=f"Error reading {self.template_paths['models_template']}: {e}")
            return
        except Exception as e:
            logging.error(f"An error occurred: {e}")

    def compress_zip(self, ifolder_path, ofolder_path=None):
        logging.debug("Compressing files into zip")
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
            prusa_model.write(model_path, encoding='utf-8', xml_declaration=True, pretty_print=True)
        except Exception as e:
            logging.error(f"An error occurred while writing Prusa object: {e}")

    def generate3mf_file(self, final_prusamodels, output_file=None):
        logging.debug("Generating 3mf file structure")
        if output_file is None:
            output_file = self.output_file
        if not output_file:
            self.status_label.config(text="Please provide an output file.")
            return
        if not final_prusamodels:
            logging.error("No models to generate 3mf file structure.")
            self.status_label.config(text="No models to generate 3mf file structure.")
            return
        try:
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
            ###--_rels/.rels---###
            # Create the relationships file
            rels_ET = ET.parse(self.template_paths['.rels_template'])
            rels_tree = rels_ET.getroot()
            for model in final_prusamodels:
                # Add a relationship for the model
                rel = ET.fromstring(f'<Relationship Target="/3D/Objects/{model[0]}" Id="rel-{model[0]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/3dmodel"/>')
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
    # app = ZipProcessorGUI(root)
    # root.mainloop()
    b3mf = f"/Users/jcacosta/Library/CloudStorage/OneDrive-TheUniversityofTexasatElPaso/3D Objects/multicolour-majoras-mask-model_files/MAJORASMASK_FULLCOLOUR_BambuStudio.3mf"
    # b3mf = f"/Users/jcacosta/Downloads/majoras_mask_nobig.zip"
    #output = f"/Users/jcacosta/Library/CloudStorage/OneDrive-TheUniversityofTexasatElPaso/3D Objects/myoutput.3mf"
    output = f"myoutput.3mf"
    a = ZipProcessorGUI(root)
    a.bambu3mf2prusa3mf(b3mf, output)
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    main()
