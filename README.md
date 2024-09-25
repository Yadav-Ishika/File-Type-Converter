# File-Type-Converter-FTC

# Description
**File Type Converter**, this `Streamlit-based` web application allows users to seamlessly convert between different file formats, including `DOCX` to `PDF`, `PDF` to `CSV`, and `image-to-PDF` conversions (`JPEG`, `PNG`, `JPG`). The application is designed to be **user-friendly**, **fast**, and **reliable**.

# Installation (Method 1)
1. Download `.zip` file for this project under `Code` option button on Repository Page
2. Extract the `.zip` file to your preferred location
3. Open any Integrated Development Environment IDE of your choice on extracted folder's location, i used `VS Code` as IDE
4. Install the neceesary libraries if not installed earlier
   Use Command :
   ```bash
   pip install -r requirements.txt
   ```

   Or Update if already installed but not updated, Use Command :
   ```bash
   pip install --upgrade -r requirements.txt
   ```
# Installation (Method 2)
1. Open any Integrated Development Environment IDE of your choice on a location where you wish to store the project, i used `VS Code` as IDE
2. Clone this `.git` repository
   Use Command :
   ```bash
   git clone https://github.com/etsryn/File-Type-Converter.git
   ```
3. Install the neceesary libraries if not installed earlier, Use Command :
   ```bash
   pip install -r requirements.txt
   ```

   Or Update if already installed but not updated, Use Command :
   ```bash
   pip install --upgrade -r requirements.txt
   ```

# Additional Requirement
1. Poppler for Windows : Link
   bash
   https://github.com/Priyanshiguptaaa/OCRLinguist.git
   
*Step 1* : Download poppler-0.68.0_x86.7z file from above linked git repository.<br />
*Step 2* : Extract the .zip file, by default in C Drive<br />
*Step 3* : Open poppler-0.68.0 then enter in bin directory, copy path to bin<br />
*Step 4* : Press windows key, then search Edit the system environment variables, click enter<br />
*Step 5* : Click on Environment Variables on bottom right corner, click<br />
*Step 6* : In System Variable section click path then Edit then New<br />
*Step 7* : Paste the copied path, click Ok, Ok then Ok<br />
*Step 8* : Restart the IDE and run the python file ftp.py   

# File Structure
📦 file-type-converter<br />
│<br />
├── 📜 app.py                 # Main Streamlit app file<br />
├── 📜 requirements.txt        # Python dependencies for the project<br />
├── 📜 README.md               # Project documentation (this file)<br />

# Usage
Since this project is a `streamlit`-based web application, so simply `py ftc.py` won't help to execute this program, instead use
   ```bash
   streamlit run ftc.py
```
## Features

- Convert `DOCX` to `PDF`
- Convert `PDF` to `DOCX`
- Convert `PDF` to `CSV`
- Convert `CSV` to `PDF`
- Convert `PDF` to `JPEG/JPG/PNG`
- Convert `JPEG/JPG/PNG` to `PDF`

# Contributing
1. You can improve the design of UI which may look more professional and easy to use
2. You may add more functionality like PDF password encryption and decryption
3. You may add functionality that allows user to rename the output file to their choice of name before downloading it

# Changelog
1. Currently working on 2nd point of Contributing section.

# Conclusion
This program is currently compiling with zero error, so you are good to go with it, **Thank You**

# License
Apache-2.0 license

# Contact me at
LinkedIn Account : https://www.linkedin.com/in/ishika-yadav-86839127a/<br />
Email At : ishikayadav11112@gmail.com
