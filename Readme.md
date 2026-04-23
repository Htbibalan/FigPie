# 🍰 FigPie:   Assemble  scientific figures as easy as pie
##### 🚧🏗️ Under construction, check for updates.🏗️🚧
#### FigPie is a lightweight, drag-and-drop GUI tool for assembling scientific figures from images, PDFs, and plots without needing heavy,expensive or overkill software like Illustrator, Affinity,  PowerPoint and such. It is designed based on my lived-experience for researchers who want fast, clean, publication-ready figures.

## 🚀 Why FigPie?
#### Many researchers and students generate plots in Python / R / MATLAB then struggle in Illustrator / PowerPoint.  FigPie removes that second step. It’s built specifically for papers , theses, presentations and scientific posters.

![Banner Image](https://github.com/Htbibalan/FigPie/blob/master/src/banner.png)

## ✨ Features

🤖 Automatic and smart panel labelling

🤖 Smart stat symbol placement

🪚 Trim, Crop and Erase unwanted objects off the plots

🧩 Flexible layout (grid, manual, alignment tools)

🔲 Add shapes: rectangle, circle, line, arrow, highlight bars

🎯 Easy precise positioning & resizing

📏 Custom and auto-adjustable canvas size

✂️ Automatic and easy white/blank-space cropping and trimming of final product

💾 Save / load projects to work later

🖼️ Export to: PNG,TIFF,SVG and PDF



## Example workflow

![stat_example_00 Image](https://github.com/Htbibalan/FigPie/blob/master/src/stat_example_00.gif)
![stat_example_01 Image](https://github.com/Htbibalan/FigPie/blob/master/src/stat_example_01.gif)
*You can select the type of stat symbol and level of statistical significance and add them on the plot by two click*


![erase_function Image](https://github.com/Htbibalan/FigPie/blob/master/src/Erase_mode.gif)
*The eraser helps you remove unwanted parts of the plots or images*

![generate_label Image](https://github.com/Htbibalan/FigPie/blob/master/src/labels.gif)
![generate_label Image](https://github.com/Htbibalan/FigPie/blob/master/src/label_regenerate.gif)
 *Generate and edit labels easily and if the order of plots/panels is changed, regenerate them automatically*


![align](https://github.com/Htbibalan/FigPie/blob/master/src/align.gif)
*Align and arrange the plots*




## 🖥️ Installation
##### Option 1: 
        
        git clone https://github.com/Htbibalan/FigPie.git
        
        cd FigPie
        
        cd Src and install the env using:

        conda env create -f environment.yml
after env is created, run:

            conda activate FigPie
Run:

            python FigPie.py
#### Option 2:
Windows users can download FigPieSetup.exe from the latest release and install it.
https://github.com/Htbibalan/FigPie/releases/latest

Security settings might stop you from downloading, or if needed, right click on the setupfile and from Properties menu check " Unblock" box.

### 🪪 Author
**Developed by Hamid Taghipourbibalan | UiT: The Arctic University of Norway, Tromsø, Norway**



### 🛠️ To-Do/ Updates needed:

~~Add crop/copy object, trim image~~

* Mac/ Linux support-test improvements

* Better snapping / alignment guides (Auto V and H needs enhancement)

~~SVG export~~

* Shortcut key fix

* Shape insert enhancement /  shape rotation fix 

* add simple manual file

* Open new canvas option / fix the override problem

* add hint on keys when mouse hovers 

~~fix Gap Y field mask~~

~~add smart stat symbol feature~~

~~warning/save before close~~ 

~~Auto plot placement fails, rewrite the function to handle smart placement based on import order~~




## Citation

If you use FigPie, please cite:

Taghipourbibalan, H. (2026). FigPie (Version 0.1.6) [Software]. Zenodo.  
https://doi.org/10.5281/zenodo.19712900