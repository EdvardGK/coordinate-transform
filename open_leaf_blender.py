"""
Blender script to import the anatomical leaf IFC and set up a nice view.
Run from terminal:  blender --python open_leaf_blender.py
Or from Blender:    File > Open > Scripting tab > paste & run
"""
import bpy
import os

IFC_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo_leaf_10x.ifc")

# Clear default scene
bpy.ops.object.select_all(action='SELECT')
bpy.ops.object.delete()

# Import IFC
bpy.ops.bim.load_project(filepath=IFC_PATH) if hasattr(bpy.ops, 'bim') else bpy.ops.import_scene.ifc(filepath=IFC_PATH)

print(f"Imported: {IFC_PATH}")
print(f"Objects: {len(bpy.data.objects)}")

# Numpad '.' to frame selected after Blender opens — do it manually
# Switch to Material Preview (Z key) to see the colors
