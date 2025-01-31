import sys
import os
import zipfile
import logging
from azure.storage.blob import BlobServiceClient
from typing import List

# Add site-packages to sys.path for external libraries
# This ensures FreeCAD's Python environment can find installed libraries like azure-storage-blob.
sys.path.append(r"C:\Users\drago\AppData\Roaming\Python\Python311\site-packages")

# Configure logging
# Provides detailed information about the execution of the plugin, useful for debugging.
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Azure Blob Storage connection details
# These need to be updated with the user's Azure connection string and container name.
CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=*******;AccountKey=******AvzgXgfzHiKtipiQcKHjDdhCFPJqeobGoIwhGVEBDTiaOM1njfevL/8HFXWIdDEjUxxSG+AStW0rM3Q==;EndpointSuffix=core.windows.net"
CONTAINER_NAME = "caddemo0"

# Define directories and file names
# These paths and names can be customized based on the user's setup.
SAVE_DIRECTORY = r"C:\Users\drago\OneDrive\Desktop\cadstuff\cad"
BOM_FILE_NAME = "BOM.csv"
ZIP_FILE_NAME = "cad_and_bom.zip"
PART_TYPE = "Part::Feature"  # Object type to include in the BOM.

def save_active_document() -> str:
    """
    Save the current FreeCAD document to the specified directory.
    Returns:
        str: The full path to the saved CAD file
    Raises:
        Exception: If no active document is found in FreeCAD.
    """
    try:
        doc = FreeCAD.ActiveDocument
        if not doc:
            raise Exception("No active document found in FreeCAD.")
        file_path = os.path.join(SAVE_DIRECTORY, f"{doc.Label}.FCStd")
        doc.saveAs(file_path)
        logger.info(f"CAD file saved: {file_path}")
        return file_path
    except Exception as e:
        logger.error(f"Failed to save active document: {e}")
        raise

def extract_bom_data() -> str:
    """
    Extract BOM data from the active FreeCAD document and save it as a CSV file.
    Returns:
        str: The full path to the saved BOM CSV file.
    """
    try:
        # Define the output BOM file path
        bom_file_path = os.path.join(SAVE_DIRECTORY, BOM_FILE_NAME)
        
        # Open the file for writing BOM data
        with open(bom_file_path, "w") as bom_file:
            # Write CSV headers
            bom_file.write("Part Name,Type,Quantity,Dimensions (LxWxH)\n")
            
            for obj in FreeCAD.ActiveDocument.Objects:
                # Skip objects with no Shape (e.g., groups, non-geometric objects)
                if not hasattr(obj, "Shape") or not obj.Shape:
                    logger.warning(f"Skipping object with no geometry: {obj.Label}")
                    continue
                
                # Include relevant object types
                if obj.TypeId in [
                    "Part::Feature",  # Basic parts
                    "Part::Box",      # Box primitives
                    "Part::Cylinder", # Cylinder primitives
                    "Part::Cut",      # Boolean cuts
                    "Part::MultiFuse", # Fused parts
                    "Part::Extrusion", # Extruded objects
                    "Part::Mirroring", # Mirrored parts
                    "App::Part",       # App containers
                    "Part::Part2DObjectPython", # Custom 2D objects
                ]:
                    # Get bounding box dimensions (fallback to "N/A" if no bounding box)
                    if hasattr(obj.Shape, "BoundBox"):
                        bbox = obj.Shape.BoundBox
                        dimensions = f"{bbox.XLength:.2f} x {bbox.YLength:.2f} x {bbox.ZLength:.2f}"
                    else:
                        dimensions = "N/A"
                    
                    # Write the object information to the BOM
                    bom_file.write(f"{obj.Label},{obj.TypeId},1,{dimensions}\n")
                else:
                    logger.warning(f"Unhandled object type: {obj.Label}, TypeId: {obj.TypeId}")
        
        logger.info(f"BOM data extracted: {bom_file_path}")
        return bom_file_path

    except Exception as e:
        logger.error(f"Failed to extract BOM data: {e}")
        raise


def compress_files(file_paths: List[str]) -> str:
    """
    Compress the CAD file and BOM data into a ZIP archive.
    Args:
        file_paths (List[str]): List of file paths to include in the ZIP file.
    Returns:
        str: The full path to the created ZIP file.
    Raises:
        Exception: If any issue occurs during compression.
    """
    try:
        zip_file_path = os.path.join(SAVE_DIRECTORY, ZIP_FILE_NAME)
        with zipfile.ZipFile(zip_file_path, "w") as zf:
            for file_path in file_paths:
                zf.write(file_path, os.path.basename(file_path))
        logger.info(f"Files compressed: {zip_file_path}")
        return zip_file_path
    except Exception as e:
        logger.error(f"Failed to compress files: {e}")
        raise

def upload_to_azure(zip_file_path: str):
    """
    Upload the ZIP file to Azure Blob Storage.
    Args:
        zip_file_path (str): Path to the ZIP file to upload.
    Raises:
        Exception: If any issue occurs during the upload process.
    """
    try:
        blob_service_client = BlobServiceClient.from_connection_string(CONNECTION_STRING)
        blob_client = blob_service_client.get_blob_client(container=CONTAINER_NAME, blob=os.path.basename(zip_file_path))
        with open(zip_file_path, "rb") as data:
            blob_client.upload_blob(data, overwrite=True)
        logger.info("File uploaded successfully to Azure!")
    except Exception as e:
        logger.error(f"Failed to upload file to Azure: {e}")
        raise

def main():
    """
    Main function to orchestrate the plugin's workflow:
    1. Save the active FreeCAD document.
    2. Extract BOM data.
    3. Compress files into a ZIP archive.
    4. Upload the ZIP archive to Azure Blob Storage.
    """
    try:
        # Save the CAD file
        cad_file = save_active_document()

        # Extract BOM data
        bom_file = extract_bom_data()

        # Compress both files
        zip_file = compress_files([cad_file, bom_file])

        # Upload to Azure
        upload_to_azure(zip_file)
        logger.info("Plugin executed successfully!")
    except Exception as e:
        logger.error(f"Error: {e}")

# Entry point for the plugin
if __name__ == "__main__":
    main()
