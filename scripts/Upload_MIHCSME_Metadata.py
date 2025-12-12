# coding=utf-8
"""
Upload_MIHCSME_Metadata.py

Uploads MIHCSME (Minimum Information about a High Content Screening
Microscopy Experiment) metadata from an Excel file to OMERO Plates or Screens.

This script parses a MIHCSME-formatted Excel file and creates Key-Value pair
annotations on OMERO objects (Screens, Plates, and Wells) with the appropriate
metadata.

-----------------------------------------------------------------------------
  Copyright (C) 2024
  This program is free software; you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation; either version 2 of the License, or
  (at your option) any later version.
  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.
  You should have received a copy of the GNU General Public License along
  with this program; if not, write to the Free Software Foundation, Inc.,
  51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
------------------------------------------------------------------------------

"""

import omero
from omero.gateway import BlitzGateway
from omero.rtypes import rstring, rlong, robject
import omero.scripts as scripts
from omero.constants.metadata import NSCLIENTMAPANNOTATION

import tempfile
import os
from pathlib import Path

# Import the MIHCSME package
try:
    from mihcsme_omero import parse_excel_to_model, upload_metadata_to_omero
    MIHCSME_AVAILABLE = True
except ImportError:
    MIHCSME_AVAILABLE = False


P_DTYPE = "Data_Type"
P_IDS = "IDs"
P_NAMESPACE = "Namespace"
P_REPLACE = "Replace existing annotations"
P_EXCEL_FILE = "MIHCSME_Excel_File"

DEFAULT_NAMESPACE = "MIHCSME"


def download_file_annotation(conn, file_ann_id):
    """
    Download a FileAnnotation to a temporary file.

    :param conn: OMERO connection
    :type conn: omero.gateway.BlitzGateway
    :param file_ann_id: ID of the FileAnnotation
    :type file_ann_id: int
    :return: Path to the downloaded temporary file
    :rtype: Path
    """
    file_ann = conn.getObject("FileAnnotation", file_ann_id)
    if file_ann is None:
        raise ValueError(f"FileAnnotation {file_ann_id} not found")

    # Get the original file
    orig_file = file_ann.getFile()
    file_name = orig_file.getName()
    file_size = orig_file.getSize()

    print(f"Downloading file: {file_name} ({file_size} bytes)")

    # Create temporary file
    tmp_dir = tempfile.mkdtemp(prefix='MIHCSME_upload_')
    tmp_path = Path(tmp_dir) / file_name

    # Download the file
    with open(tmp_path, 'wb') as f:
        for chunk in orig_file.getFileInChunks():
            f.write(chunk)

    print(f"Downloaded to: {tmp_path}")
    return tmp_path


def main_loop(conn, script_params):
    """
    Main processing loop for uploading MIHCSME metadata.

    :param conn: OMERO connection
    :type conn: omero.gateway.BlitzGateway
    :param script_params: Script parameters from client
    :type script_params: dict
    :return: Summary message and result object
    :rtype: tuple
    """
    if not MIHCSME_AVAILABLE:
        raise ImportError(
            "The mihcsme_omero package is not installed on this OMERO server. "
            "Please install it using: pip install mihcsme-omero"
        )

    target_type = script_params[P_DTYPE]
    target_ids = script_params[P_IDS]
    namespace = script_params[P_NAMESPACE]
    replace = script_params[P_REPLACE]
    file_ann_id = script_params[P_EXCEL_FILE]

    # Download the Excel file
    tmp_file_path = None
    try:
        tmp_file_path = download_file_annotation(conn, file_ann_id)

        # Parse the Excel file
        print(f"\nParsing MIHCSME Excel file: {tmp_file_path.name}")
        metadata = parse_excel_to_model(tmp_file_path)
        print(f"✓ Successfully parsed metadata")
        print(f"  - Investigation groups: {len(metadata.investigation_information.groups)}")
        print(f"  - Study information entries: {len(metadata.study_information.metadata)}")
        print(f"  - Assay information entries: {len(metadata.assay_information.metadata)}")
        print(f"  - Assay conditions (wells): {len(metadata.assay_conditions)}")

        # Process each target object
        results = []
        for target_id in target_ids:
            target_obj = conn.getObject(target_type, target_id)
            if target_obj is None:
                print(f"\n✗ {target_type} {target_id} not found, skipping")
                continue

            print(f"\n{'='*60}")
            print(f"Processing {target_type} {target_id}: {target_obj.getName()}")
            print(f"{'='*60}")

            # Upload metadata
            result = upload_metadata_to_omero(
                conn=conn,
                metadata=metadata,
                target_type=target_type,
                target_id=target_id,
                namespace=namespace,
                replace=replace
            )

            results.append(result)

            # Print summary
            print(f"\n✓ Upload completed:")
            print(f"  - Status: {result['status']}")
            print(f"  - Screen annotation: {'✓' if result.get('screen_annotated') else '✗'}")
            print(f"  - Plates annotated: {result.get('plates_annotated', 0)}")
            print(f"  - Wells succeeded: {result['wells_succeeded']}/{result['wells_total']}")
            if result['wells_failed'] > 0:
                print(f"  - Wells failed: {result['wells_failed']}")
                if result.get('errors'):
                    print(f"  - Errors: {result['errors'][:3]}")  # Show first 3 errors

        # Create summary message
        total_wells = sum(r['wells_total'] for r in results)
        successful_wells = sum(r['wells_succeeded'] for r in results)
        failed_wells = sum(r['wells_failed'] for r in results)

        message = (
            f"MIHCSME metadata upload completed.\n"
            f"Processed {len(target_ids)} {target_type}(s).\n"
            f"Wells: {successful_wells}/{total_wells} succeeded"
        )
        if failed_wells > 0:
            message += f", {failed_wells} failed"

        # Return the first target as result object
        result_obj = conn.getObject(target_type, target_ids[0])

        return message, result_obj

    finally:
        # Clean up temporary file
        if tmp_file_path and tmp_file_path.exists():
            tmp_file_path.unlink()
            tmp_file_path.parent.rmdir()
            print(f"\nCleaned up temporary file: {tmp_file_path}")


def run_script():
    """
    Entry point for the OMERO script.
    """

    data_types = [rstring("Screen"), rstring("Plate")]

    client = scripts.client(
        'Upload MIHCSME Metadata',
        """
    Uploads MIHCSME metadata from an Excel file to OMERO Screens or Plates.

    This script parses a MIHCSME-formatted Excel file and creates Key-Value
    pair annotations on the selected Screen/Plate and associated Wells.

    Requirements:
    - The mihcsme_omero package must be installed on the OMERO server
    - Upload a MIHCSME Excel file as a FileAnnotation first
    - Select the target Screen or Plate to annotate

    The script will:
    1. Download and parse the Excel file
    2. Create Investigation/Study/Assay annotations on the Screen/Plate
    3. Create per-well condition annotations on each Well

    Default namespace: MIHCSME
        """,

        scripts.String(
            P_DTYPE, optional=False, grouping="1",
            description="Type of object to annotate (Screen or Plate)",
            values=data_types, default="Plate"),

        scripts.List(
            P_IDS, optional=False, grouping="1.1",
            description="IDs of Screen(s) or Plate(s) to annotate"
        ).ofType(rlong(0)),

        scripts.Long(
            P_EXCEL_FILE, optional=False, grouping="2",
            description="FileAnnotation ID of the MIHCSME Excel file to upload"
        ),

        scripts.String(
            P_NAMESPACE, optional=True, grouping="3",
            description="Namespace for the annotations",
            default=DEFAULT_NAMESPACE),

        scripts.Bool(
            P_REPLACE, optional=True, grouping="3.1",
            description="Replace existing annotations in this namespace",
            default=False),

        authors=["Maarten Paul"],
        institutions=["Your Institution"],
        contact="https://forum.image.sc/tag/omero",
        version="1.0.0",
    )

    try:
        # Get parameters
        params = {}
        for key in client.getInputKeys():
            if client.getInput(key):
                params[key] = client.getInput(key, unwrap=True)

        # Set defaults
        if P_NAMESPACE not in params:
            params[P_NAMESPACE] = DEFAULT_NAMESPACE
        if P_REPLACE not in params:
            params[P_REPLACE] = False

        print("Input parameters:")
        for k, v in params.items():
            print(f"  - {k}: {v}")
        print("\n" + "="*60 + "\n")

        # Connect to OMERO
        conn = BlitzGateway(client_obj=client)

        # Run main loop
        message, result_obj = main_loop(conn, params)

        # Set outputs
        client.setOutput("Message", rstring(message))
        if result_obj is not None:
            client.setOutput("Result", robject(result_obj._obj))

    except ImportError as err:
        error_msg = str(err)
        print(f"ERROR: {error_msg}")
        client.setOutput("Message", rstring(f"ERROR: {error_msg}"))
        raise

    except Exception as err:
        error_msg = f"Error during MIHCSME upload: {str(err)}"
        print(f"ERROR: {error_msg}")
        client.setOutput("Message", rstring(error_msg))
        raise

    finally:
        client.closeSession()


if __name__ == "__main__":
    run_script()
