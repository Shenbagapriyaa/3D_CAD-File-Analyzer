import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import numpy as np
import trimesh
from stl import mesh as stl_mesh
import cadquery as cq
import openpyxl
import itertools

processed_data = []

def calculate_overhang_area(tri_mesh, build_direction=[0, 0, -1], angle_threshold=45):
    face_normals = tri_mesh.face_normals
    face_areas = tri_mesh.area_faces

    build_direction = np.array(build_direction)
    build_direction = build_direction / np.linalg.norm(build_direction)

    dot_products = face_normals @ build_direction
    dot_products = np.clip(dot_products, -1.0, 1.0)
    angles = np.degrees(np.arccos(dot_products))

    overhang_faces = angles < angle_threshold
    return round(face_areas[overhang_faces].sum(), 2)

def process_file(file_path):
    ext = os.path.splitext(file_path)[-1].lower()
    try:
        tri_mesh = None
        if ext == '.stl':
            your_mesh = stl_mesh.Mesh.from_file(file_path)
            tri_mesh = trimesh.load(file_path, force='mesh')
        elif ext in ['.obj', '.ply']:
            tri_mesh = trimesh.load(file_path, force='mesh')
        elif ext in ['.step', '.stp', '.iges', '.igs']:
            cq_obj = cq.importers.importStep(file_path) if ext in ['.step', '.stp'] else cq.importers.importIges(file_path)
            shape = cq_obj.val() if hasattr(cq_obj, 'val') else cq_obj
            tess = shape.tessellate(0.1)
            vertices_raw, faces = tess
            vertices = np.array([(v.x, v.y, v.z) for v in vertices_raw])
            faces = np.array(faces)
            tri_mesh = trimesh.Trimesh(vertices=vertices, faces=faces)
        else:
            return "Unsupported file format", None

        size_x, size_y, size_z = tri_mesh.bounding_box.extents
        volume = round(tri_mesh.volume, 2)
        surface_area = round(tri_mesh.area, 2)
        overhang_area = calculate_overhang_area(tri_mesh)

        machine_limits = {
            "EOS M280/290": (250, 250, 315),
            "EOS M400": (400, 400, 400),
            "SLM 500": (500, 250, 315),
        }

        orientations = list(itertools.permutations([size_x, size_y, size_z]))
        fitting_machines = []
        for name, (mx, my, mz) in machine_limits.items():
            for ox, oy, oz in orientations:
                if ox <= mx and oy <= my and oz <= mz:
                    fitting_machines.append(name)
                    break

        result = f"""
File: {os.path.basename(file_path)}
Volume: {volume} mm³
Surface Area: {surface_area} mm²
Overhang Area (<45°): {overhang_area} mm²
Dimensions: {round(size_x, 2)} × {round(size_y, 2)} × {round(size_z, 2)} mm
Best Fit Machine: {', '.join(fitting_machines) if fitting_machines else 'None'}
"""

        processed_data.append([
            os.path.basename(file_path),
            round(size_x, 2),
            round(size_y, 2),
            round(size_z, 2),
            volume,
            surface_area,
            overhang_area,
            ', '.join(fitting_machines) if fitting_machines else 'None'
        ])
        return result, None

    except Exception as e:
        
        print(f"DEBUG ERROR: {e}") 
        
        return None, str(e)

def open_files():
    file_paths = filedialog.askopenfilenames(
        title="Select 3D Files",
        filetypes=[("3D files", "*.stl *.obj *.ply *.step *.stp *.iges *.igs")]
    )

    if file_paths:
        output_text.delete("1.0", tk.END)
        window.update_idletasks()

        processed_data.clear()
        progress["maximum"] = len(file_paths)
        progress["value"] = 0

        for idx, path in enumerate(file_paths, start=1):
            result, error = process_file(path)
            if error:
                output_text.insert(tk.END, f"{os.path.basename(path)} - ERROR: {error}\n\n")
            else:
                output_text.insert(tk.END, result + "\n")
            progress["value"] = idx
            window.update_idletasks()

        output_text.insert(tk.END, "All files processed.\n")

def export_to_excel():
    if not processed_data:
        messagebox.showwarning("No Data", "Please analyze files before exporting.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if not save_path:
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orientation Assist Results"

    headers = ["File", "X (mm)", "Y (mm)", "Z (mm)", "Volume (mm³)", "Surface Area (mm²)", "Overhang Area (mm²)", "Best Fit Machine"]
    ws.append(headers)

    for row in processed_data:
        ws.append(row)

    wb.save(save_path)
    messagebox.showinfo("Export Successful", f"Data exported to {save_path}")

# -------------------- GUI --------------------
window = tk.Tk()
window.title("Orientation Assist - 3D AM Machine Checker")
window.geometry("750x520")
frame = tk.Frame(window)
frame.pack(pady=20)
btn_select = tk.Button(frame, text="Select 3D File(s)", command=open_files, font=("Arial", 13))
btn_select.pack(side=tk.LEFT, padx=10)
btn_export = tk.Button(frame, text="Export to Excel", command=export_to_excel, font=("Arial", 13))
btn_export.pack(side=tk.LEFT, padx=10)
progress = ttk.Progressbar(window, orient="horizontal", length=650, mode="determinate")
progress.pack(pady=5)
output_text = tk.Text(window, wrap=tk.WORD, font=("Courier", 11))
output_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
window.mainloop()
