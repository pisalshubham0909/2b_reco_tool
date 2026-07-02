# DEPLOYMENT GUIDE: GSTR-2B Reconciliation Tool

This guide explains how to host and publish your Streamlit reconciliation tool online so your team or clients can access it from anywhere.

---

## 🔒 Important: Data Security & Privacy
Since this tool processes sensitive corporate financial records (GSTINs, invoices, and accounting registers):
* **No Database Logging**: Our application does not store uploaded files on a server or write them to a database. Uploaded files are parsed completely in-memory and discarded as soon as the user closes the browser session.
* **Hosting Security**: When publishing, choose a deployment option that aligns with your company's IT data protection policies.

---

## Option 1: Streamlit Community Cloud (Easiest & Free)
Streamlit offers free hosting directly from GitHub.

### Steps:
1. **Create a GitHub Account**: If you don't have one, sign up at [github.com](https://github.com).
2. **Create a Repository**:
   * Create a new repository on your GitHub account page. You can set it to **Private** so that only you and your team can see the code.
   * **Choose one of these 3 methods to upload the code**:
     * **Method A: Direct Website Drag & Drop (Easiest)**:
       On your new GitHub repository page, click the **"uploading an existing file"** link. Drag and drop all the files from this folder (`app.py`, `parser.py`, `reconciliation.py`, `requirements.txt`, `.gitignore`, and `.streamlit/config.toml`) directly into your browser window, then click **"Commit changes"**.
     * **Method B: Using GitHub Desktop App (Recommended for Windows)**:
       1. Download and install [GitHub Desktop](https://desktop.github.com/).
       2. Open the app, log in to your GitHub account, and go to **File > Add Local Repository...**
       3. Select your project folder: `C:\Users\spisa\OneDrive\Automation folder\2B Reco Tool`
       4. GitHub Desktop will ask if you want to initialize it as a repository; click **Yes**.
       5. Enter a commit message (e.g., "Initial commit") and click **"Commit to main"**.
       6. Click **"Publish Repository"**, choose Private/Public, and upload it.
     * **Method C: Using Git Command Line**:
       If you install Git from [git-scm.com](https://git-scm.com), open a terminal in this folder and run:
       ```bash
       git init
       git add .
       git commit -m "Initial commit of GSTR-2B Reco Tool"
       git branch -M main
       git remote add origin <YOUR_GITHUB_REPOSITORY_URL>
       git push -u origin main
       ```
3. **Connect to Streamlit Cloud**:
   * Visit [share.streamlit.io](https://share.streamlit.io) and log in using your GitHub account.
   * Click **"Create app"** (or **"New app"**).
4. **Configure & Deploy**:
   * Select your **Repository**, **Branch** (usually `main` or `master`), and set **Main file path** to `app.py`.
   * Click **"Deploy!"**.
   * Streamlit will compile the `requirements.txt` dependencies and launch your live website (usually within 1–2 minutes).

---

## Option 2: Render or Hugging Face Spaces (Free Cloud Alternatives)
These platforms allow easy web deployment with automatic GitHub syncs.

### Using Hugging Face Spaces:
1. Create an account on [huggingface.co](https://huggingface.co).
2. Go to **Spaces** -> **Create New Space**.
3. Set Space SDK to **Streamlit**.
4. Set visibility to **Public** or **Private**.
5. Upload your files via git or drag-and-drop, and Hugging Face will host the app automatically.

### Using Render:
1. Create a free account on [render.com](https://render.com).
2. Link your GitHub account and select your project repository.
3. Create a **Web Service**:
   * **Build Command**: `pip install -r requirements.txt`
   * **Start Command**: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
4. Deploy the service. Render will provide a free public URL (e.g. `your-app.onrender.com`).

---

## Option 3: AWS / Azure / Google Cloud VM (Recommended for Corporate Data)
If you require maximum security and need to deploy behind a private virtual network (VPN), host it on a private virtual machine.

### Setup Steps (Ubuntu Linux Server Example):
1. **Launch a VM instance** (AWS EC2, Azure VM, etc.) on your cloud console.
2. **Install Python & Git**:
   ```bash
   sudo apt update
   sudo apt install python3-pip python3-venv git -y
   ```
3. **Clone your code**:
   ```bash
   git clone <your-private-repo-url> /var/www/reco-tool
   cd /var/www/reco-tool
   ```
4. **Set up a Virtual Environment**:
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```
5. **Configure systemd to run Streamlit in the background**:
   Create a service configuration file:
   ```bash
   sudo nano /etc/systemd/system/streamlit.service
   ```
   Paste the following:
   ```ini
   [Unit]
   Description=Streamlit GSTR-2B Reco Service
   After=network.target

   [Service]
   User=ubuntu
   WorkingDirectory=/var/www/reco-tool
   ExecStart=/var/www/reco-tool/venv/bin/streamlit run app.py --server.port 80 --server.address 0.0.0.0
   Restart=always

   [Install]
   WantedBy=multi-user.target
   ```
6. **Start and enable the service**:
   ```bash
   sudo systemctl daemon-reload
   sudo systemctl start streamlit
   sudo systemctl enable streamlit
   ```
   *Your site will now run continuously in the background on port `80` (accessible via the VM's public IP or assigned DNS domain).*
