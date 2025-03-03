/* global document, Office, Word */
import axios from 'axios';
// import { jwtdecode } from "jwt-decode";
import { jwtDecode, JwtPayload } from 'jwt-decode'



let currentPlaceholders = {};

// Event listeners for buttons in the task pane
// document.getElementById("loadTemplate").addEventListener("click", loadTemplate);
document.getElementById("applyEditedPlaceholders").addEventListener("click", updatePlaceholders);
document.getElementById("enableTrackChanges").addEventListener("click", enableTrackChanges);
document.getElementById("generateAIChangesWithContext").addEventListener("click", generateAIChangesWithContext);

document.getElementById("disableAllChanges").addEventListener("click", disableAllChanges);
document.getElementById("listTrackedChanges").addEventListener("click", listTrackedChanges);
document.getElementById("acceptSelectedChanges").addEventListener("click", acceptSelectedTrackedChanges);
document.getElementById("rejectSelectedChanges").addEventListener("click", rejectSelectedTrackedChanges);
document.getElementById("acceptAllChanges").addEventListener("click", acceptAllTrackedChanges);
document.getElementById("rejectAllChanges").addEventListener("click", rejectAllTrackedChanges);
document.getElementById("templatesBtn").addEventListener("click", getTemplatesAndPopulateDropdown);
document.getElementById("openTemplateBtn").addEventListener("click", fetchAndOpenTemplate);
// document.getElementById("loginButton").addEventListener("click", initializeTaskPane);


// // document.getElementById("saveAdminSettings").addEventListener("click", saveAdminSettings);


Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Attach event listeners
        document.getElementById("docxFileInput").addEventListener("change", handleDocxUpload);
        document.getElementById("applyEditedPlaceholders").addEventListener("click", mergeAndInsertTemplate);
        storeCompanyDetailsInSession();


    }
});


Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      // Attach main button click handler
      document.getElementById("getIDToken").addEventListener("click", handleAdminFlow);
      document.getElementById("saveAdminSettings").addEventListener("click", saveConfig);
    }
  });
  
  async function handleAdminFlow() {
    try {
    //   let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
    //     allowSignInPrompt: true,
    //   });
    //   let userToken = jwtDecode(userTokenEncoded, { complete: true });
    //   jwt.default.decode(token, { complete: true });
    //   document.getElementById("userInfo").innerHTML =
    //     "name: " +
    //     userToken.name +
    //     "<br>email: " +
    //     userToken.preferred_username +
    //     "<br>id: " +
    //     userToken.oid;
    //   insertDebugMessage(userToken);
    //   insertDebugMessage(userToken.preferred_username)

      const storedDetails = sessionStorage.getItem('companyDetails');

        if (storedDetails) {
        const companyDetails = JSON.parse(storedDetails);
        insertDebugMessage(`Retrieved company details:", ${companyDetails}`);
        insertDebugMessage(`Retrieved company email:", ${companyDetails.email}`);
      
        insertDebugMessage(`Retrieved company name:", ${companyDetails.domain}`);

        insertDebugMessage(`Retrieved company email:", ${companyDetails.companyName}`);


        const email = companyDetails.email;
        const company = email.split('@')[1].split('.')[0];
         
        // 3. Check admin status with backend
        const { is_admin, config } = await checkAdminStatus(email);
     //    const isAdmin  = await checkAdminStatus(email);
        insertDebugMessage(`printingt he admin, ${is_admin}`)
        
        // 4. Toggle admin UI
        const adminSection = document.getElementById('adminSection');
        adminSection.style.display = is_admin ? 'block' : 'none';
        
        // 5. If admin, load company config
        if (is_admin && config) {
            document.getElementById('apiKey').value = config.openai_key || '';
            document.getElementById('onedriveLink').value = config.onedrive_link || '';
        }else{
            const debugContainer = document.getElementById("debugMessages");
            debugContainer.style.display = 'block'; // Show on error
            debugContainer.innerHTML = "You do not have Admin Permission"
        }


        // Use companyDetails.email, companyDetails.domain, or companyDetails.companyName as needed.
        }

    //    // 2. Extract company name from email
    //    const email = userToken.preferred_username
      

    } catch (error) {
    //   console.log(error);
    //   document.getElementById("debugMessages").innerHTML = error;
    //   debugContainer.style.display = 'block';
        const debugContainer = document.getElementById("debugMessages");
        debugContainer.style.display = 'block'; // Show on error
        debugContainer.innerHTML = "You do not have Admin Permission"

    }
  }
  

async function storeCompanyDetailsInSession() {
    try {
      // Retrieve the access token from Office and decode it to get user details
      const token = await Office.auth.getAccessToken();
      const decodedToken = jwtDecode(token);
      const email = decodedToken.preferred_username; // e.g. "user@company.onmicrosoft.com"
  
      // Extract the domain (everything after the '@')
      const domain = email.split('@')[1];
  
      // Derive a simple company name from the domain (e.g. "company" from "company.onmicrosoft.com")
      const companyName = domain.split('.')[0];
  
      // Create an object with the user and company details
      const companyDetails = {
        email,
        domain,
        companyName
      };
  
      // Store the details in session storage as a JSON string
      sessionStorage.setItem('companyDetails', JSON.stringify(companyDetails));
  
      console.log("Stored company details in session:", companyDetails);
    } catch (error) {
      console.error("Error storing company details in session:", error);
    }
}

  
  function getCompanyFromEmail(email) {
    // user@company.domain → "company"
    const domainPart = email.split('@')[1];
    return domainPart.split('.')[0];
}


// Update saveConfig to match backend expectations
async function saveConfig() {
    try {
        const email = jwtDecode(await Office.auth.getAccessToken()).preferred_username;
        const domain = email.split('@')[1]; // Match backend's domain extraction
        
        await axios.post(
            'https://a169-2607-fea8-fc01-7009-d074-5ac-b353-5829.ngrok-free.app/api/save-config',
            {
                domain, // Send full domain instead of company name
                openai_key: document.getElementById('apiKey').value,
                onedrive_link: document.getElementById('onedriveLink').value
            },
            {
                headers: {
                    "ngrok-skip-browser-warning": "true",
                    "Content-Type": "application/json"
                }
            }
        );
        
        alert('Settings saved successfully!');
    } catch (error) {
        console.error('Save failed:', error);
        alert(`Failed to save settings: ${error.response?.data?.detail || error.message}`);
    }
}


//  // Step 1: Fetch the .docx file from backend
//  const response = await axios.get(`https://bca5-142-126-182-191.ngrok-free.app/api/templates/${selectedTemplateId}`, {
//     headers: { "ngrok-skip-browser-warning": "true" },
//     responseType: "arraybuffer" // ⚠️ Change response type to arraybuffer
// });

// Simplified backend calls
async function checkAdminStatus(email) {
    try {
        const response = await axios.post('https://a169-2607-fea8-fc01-7009-d074-5ac-b353-5829.ngrok-free.app/api/check-admin', 
        { email },
        { headers:
            { "ngrok-skip-browser-warning": "true",
                "Content-Type": "application/json",
            }
        }
    );
        insertDebugMessage(`Admin check result:', ${response.data.is_admin}`);

        return {
            is_admin: response.data.is_admin,
            config: response.data.config
        };
    } catch (error) {
        insertDebugMessage(`Admin check failed:', ${error}`);
        return { isAdmin: false };
    }
}

// Function to handle DOCX upload
function handleDocxUpload(event) {
    const file = event.target.files[0];
    const statusEl = document.getElementById("uploadStatus");

    if (!file) {
        statusEl.textContent = "No file selected.";
        return;
    }
    if (file.type !== "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
        statusEl.textContent = "Please upload a valid .docx file.";
        return;
    }

    statusEl.textContent = `Uploading ${file.name}...`;

    const reader = new FileReader();
    reader.onload = async function (e) { // Marked as async
        try {
            // Read file as binary string
            const content = e.target.result;

            // Insert the uploaded DOCX file into the Word document
            await Word.run(async (context) => {
                context.document.body.clear();
                context.document.body.insertFileFromBase64(arrayBufferToBase64(content), Word.InsertLocation.start);
                await context.sync();

                // Extract placeholders from the document
                const placeholders = await extractPlaceholdersFromDocument(context);
                insertDebugMessage("Extracted placeholders: " + JSON.stringify(placeholders));

                // Generate dynamic form fields for the placeholders
                generateEditFields(placeholders);

                statusEl.textContent = "File uploaded and placeholders extracted successfully.";
            });
        } catch (error) {
            console.error("Error processing DOCX file:", error);
            statusEl.textContent = "Error processing DOCX file: " + error.message;
        }
    };
    reader.onerror = function (e) {
        console.error("FileReader error:", e);
        statusEl.textContent = "Error reading file.";
    };

    // Read as ArrayBuffer
    reader.readAsArrayBuffer(file);
}


// Generate dynamic form fields for each placeholder
function generateEditFields(placeholders) {
    const container = document.getElementById("editPlaceholderFields");
    container.innerHTML = ""; // Clear previous fields
    if (!placeholders.length) {
        container.innerHTML = "<p>No placeholders found.</p>";
        return;
    }
    placeholders.forEach((ph) => {
        const div = document.createElement("div");
        div.className = "form-group";

        const label = document.createElement("label");
        label.setAttribute("for", `placeholder-${ph}`);
        label.textContent = ph + ": ";

        const input = document.createElement("input");
        input.type = "text";
        input.id = `placeholder-${ph}`;
        input.placeholder = `Enter ${ph}`;

        div.appendChild(label);
        div.appendChild(input);
        container.appendChild(div);
    });
}

// Function to replace placeholders in the document with user input
async function mergeAndInsertTemplate() {
    try {
        // Replace placeholders in the document
        await Word.run(async (context) => {
            // Extract placeholders from the document
            const placeholders = await extractPlaceholdersFromDocument(context);
            insertDebugMessage("Extracted placeholders: " + JSON.stringify(placeholders));

            // Gather user input values for each placeholder
            const data = {};
            placeholders.forEach((ph) => {
                const input = document.getElementById(`placeholder-${ph}`);
                data[ph] = input ? input.value : "";
            });
            insertDebugMessage("User data: " + JSON.stringify(data));

            // Iterate through sections (headers, body, and footers)
            const sections = context.document.sections;
            sections.load("items");
            await context.sync();

            for (const section of sections.items) {
                const parts = [
                    section.getHeader("Primary"),
                    section.body,
                    section.getFooter("Primary"),
                ];

                for (const part of parts) {
                    part.load("text");
                    await context.sync();

                    let content = part.text || "";

                    // Replace placeholders with user input
                    for (let key in data) {
                        const regex = new RegExp(`{{${key}}}`, "g");
                        content = content.replace(regex, data[key]);
                    }

                    // Clear the part and insert the updated content
                    part.clear();
                    part.insertText(content, Word.InsertLocation.replace);
                }
            }

            await context.sync();
            insertDebugMessage("Placeholders updated successfully!");
        });
    } catch (error) {
        console.error("Error during merge and render:", error);
        insertDebugMessage("Error during merge: " + error.message);
    }
}

// Helper function to extract placeholders from the document
async function extractPlaceholdersFromDocument(context) {
    try {
        const body = context.document.body;
        if (!body) {
            insertDebugMessage("Error: context.document.body is undefined in extractPlaceholdersFromDocument.");
            return [];
        }

        body.load("text"); // Load the text content of the document
        await context.sync();
        insertDebugMessage("insert if it has recahged this ")

        // Regex to find placeholders like {{Placeholder}}
        const placeholderRegex = /\{\{(.*?)\}\}/g;
        const matches = body.text.match(placeholderRegex);

        if (!matches) {
            return [];
        }

        // Extract unique placeholders
        const uniquePlaceholders = [...new Set(matches)];
        return uniquePlaceholders.map((ph) => ph.replace(/\{\{|\}\}/g, ""));
    } catch (error) {
        console.error("Error in extractPlaceholdersFromDocument:", error);
        insertDebugMessage("Error in extractPlaceholdersFromDocument: " + error.message);
        return [];
    }
}



// const templates = {
//     template1: {
//         header: "Header: {{Name}} - {{Date}}",
//         body: "Dear {{Name}},\nToday is {{Date}}.\nBest regards,\nYour Company.",
//         footer: "Footer: Sincerely, {{Signature}}",
//         placeholders: ["Name", "Date", "Signature"],
//     },
//     template2: {
//         header: "Appointment Reminder: {{Name}} - {{Date}}",
//         body: "Hello {{Name}},\nYour appointment is scheduled for {{Date}}.\nThank you!",
//         footer: "Footer: Best regards,\n{{Signature}}",
//         placeholders: ["Name", "Date", "AppointmentTime", "Signature"],
//     },
// };

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        // Auto-fetch templates as soon as the task pane is ready
        getTemplatesAndPopulateDropdown();
    }
});


// async function loadTemplate() {
//     const templateKey = document.getElementById("templateSelector").value;
//     const template = templates[templateKey];

//     await Word.run(async (context) => {
//         const sections = context.document.sections;
//         sections.load("items");
//         await context.sync();

//         sections.items.forEach((section) => {
//             const header = section.getHeader("Primary");
//             header.clear();
//             header.insertText(template.header, Word.InsertLocation.replace);

//             const body = section.body;
//             body.clear();
//             body.insertText(template.body, Word.InsertLocation.start);

//             const footer = section.getFooter("Primary");
//             footer.clear();
//             footer.insertText(template.footer, Word.InsertLocation.replace);
//         });

//         await context.sync();
//     });

//     generateEditFields(template.placeholders);
// }




async function getTemplatesAndPopulateDropdown() {
    try {

        const storedDetails = sessionStorage.getItem('companyDetails');
            
        if (!storedDetails) {
            insertDebugMessage("Please login to use this add-in");
            document.getElementById("userInfo").innerHTML = 
                "Please login to use this feature";
            return; // Stop execution if no company details
        }

        const companyDetails = JSON.parse(storedDetails);
        const company = companyDetails.companyName; // Get company name

        // Step 2: Prepare the payload
        const payload = {
            companyName : company// Send raw JSON, not stringified
        };

        // Step 1: Fetch data from the backend endpoint with Axios
        //   const response = await axios.get("https://bca5-142-126-182-191.ngrok-free.app/api/templates");

        const response = await axios.post("https://a169-2607-fea8-fc01-7009-d074-5ac-b353-5829.ngrok-free.app/api/templates",payload, {
            headers: {
            "ngrok-skip-browser-warning": "true"
            }
        });
  
      await insertDebugMessage(` please print thus repsosnise ${JSON.stringify(response)}`);
      // Step 2: Parse the JSON response (Axios auto-parses JSON by default)
      const templates = response.data.templates || [];
  
      // Step 3: Find the dropdown element
      const dropdown = document.getElementById("templatesDropdown");
      dropdown.innerHTML = ""; // Clear existing items
      templates.forEach((tpl) => {
        const option = document.createElement("option");
        option.value = tpl.id;
        option.textContent = tpl.name;
        dropdown.appendChild(option);
      });
    //   await insertDebugMessage(` lets print the dorpdown ${dropdown}`);
      console.log("Successfully populated the templates dropdown!");
    } catch (error) {
      console.error("Error fetching templates:", error);
    }
}


async function fetchAndOpenTemplate() {
    try {
        const selectedTemplateId = document.getElementById("templatesDropdown").value;
        if (!selectedTemplateId) {
            console.error("No template selected.");
            return;
        }

        const storedDetails = sessionStorage.getItem('companyDetails');
            
        if (!storedDetails) {
            insertDebugMessage("Please login to use this add-in");
            document.getElementById("userInfo").innerHTML = 
                "Please login to use this feature";
            return; // Stop execution if no company details
        }

        const companyDetails = JSON.parse(storedDetails);
        const company = companyDetails.companyName; // Get company name

        // Step 2: Prepare the payload
        const payload = {
            companyName : company// Send raw JSON, not stringified
        };


        console.log(`Fetching template with ID: ${selectedTemplateId}`);

        // Step 1: Fetch the .docx file from backend
        const response = await axios.post(`https://a169-2607-fea8-fc01-7009-d074-5ac-b353-5829.ngrok-free.app/api/templates/${selectedTemplateId}`, payload, {
            headers: { "ngrok-skip-browser-warning": "true" },
            responseType: "arraybuffer" // ⚠️ Change response type to arraybuffer
        });

        insertDebugMessage(`Template file received from API:", ${response}`);

        // Step 2: Convert ArrayBuffer to Base64
        const { fileData, placeholders } = response.data;
        const base64Data = arrayBufferToBase64(response.data);

        insertDebugMessage("Converted file to Base64, attempting to insert into Word...");

        // Step 3: Insert into Word
        await Word.run(async (context) => {
            const body = context.document.body;
            body.clear();
            body.insertFileFromBase64(base64Data, Word.InsertLocation.start);
            await context.sync();
        });

        insertDebugMessage("✅ Template successfully inserted into Word!");
        generateEditFields(placeholders);

    } catch (error) {
        insertDebugMessage(`❌ Error fetching template:", ${error}`);
    }
}


async function updatePlaceholders() {
    const templateKey = document.getElementById("templateSelector").value;
    const template = templates[templateKey];

    template.placeholders.forEach((placeholder) => {
        const input = document.getElementById(`edit-${placeholder}`);
        if (input) {
            currentPlaceholders[placeholder] = input.value || "";
        }
    });

    await Word.run(async (context) => {
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();

        for (const section of sections.items) {
            const parts = [
                section.getHeader("Primary"),
                section.body,
                section.getFooter("Primary"),
            ];

            for (const part of parts) {
                part.load("text");
                await context.sync();

                let content = part.text || "";

                for (let key in currentPlaceholders) {
                    const currentValue = currentPlaceholders[key];
                    const restoreRegex = new RegExp(currentValue, "g");
                    content = content.replace(restoreRegex, `{{${key}}}`);
                }

                for (let key in currentPlaceholders) {
                    const regex = new RegExp(`{{${key}}}`, "g");
                    content = content.replace(regex, currentPlaceholders[key]);
                }

                part.clear();
                part.insertText(content, Word.InsertLocation.replace);
            }
        }

        await context.sync();
    });
}

async function enableTrackChanges() {
    await Word.run(async (context) => {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await context.sync();
    });
}



async function listTrackedChanges() {
    await Word.run(async (context) => {
        const body = context.document.body;
        const trackedChanges = body.getTrackedChanges();
        trackedChanges.load("items");
        await context.sync();

        const changesListContainer = document.getElementById("trackedChangesList");
        changesListContainer.innerHTML = "";
        changesListContainer.style.display = 'block'; 

        if (trackedChanges.items.length === 0) {
            changesListContainer.innerHTML = "<p>No tracked changes found.</p>";
            return;
        }

        for (let i = 0; i < trackedChanges.items.length; i++) {
            const change = trackedChanges.items[i];
            change.load("type");
            const range = change.getRange();
            range.load("text");
            await context.sync();

            let changeText = range.text || "[No visible text or formatting change]";

            const listItem = document.createElement("div");
            listItem.className = "change-item";

            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.dataset.changeId = i;

            const label = document.createElement("label");
            label.textContent = `Change ${i + 1}: ${changeText}`;

            listItem.appendChild(checkbox);
            listItem.appendChild(label);
            changesListContainer.appendChild(listItem);
        }
        await context.sync();
    });
}

async function acceptSelectedTrackedChanges() {
    await Word.run(async (context) => {
        const trackedChanges = context.document.body.getTrackedChanges();
        trackedChanges.load("items");
        await context.sync();

        const selectedChanges = Array.from(
            document.querySelectorAll("#trackedChangesList input[type='checkbox']:checked")
        );

        if (selectedChanges.length === 0) return;

        selectedChanges.forEach((checkbox) => {
            const changeIndex = parseInt(checkbox.dataset.changeId, 10);
            trackedChanges.items[changeIndex].accept();
        });

        await context.sync();
    });
}

async function rejectSelectedTrackedChanges() {
    await Word.run(async (context) => {
        const trackedChanges = context.document.body.getTrackedChanges();
        trackedChanges.load("items");
        await context.sync();

        const selectedChanges = Array.from(
            document.querySelectorAll("#trackedChangesList input[type='checkbox']:checked")
        );

        if (selectedChanges.length === 0) return;

        selectedChanges.forEach((checkbox) => {
            const changeIndex = parseInt(checkbox.dataset.changeId, 10);
            trackedChanges.items[changeIndex].reject();
        });

        await context.sync();
    });
}

async function acceptAllTrackedChanges() {
    await Word.run(async (context) => {
        const trackedChanges = context.document.body.getTrackedChanges();
        trackedChanges.load("items");
        await context.sync();

        trackedChanges.items.forEach((change) => change.accept());
        await context.sync();
    });
}

async function rejectAllTrackedChanges() {
    await Word.run(async (context) => {
        const trackedChanges = context.document.body.getTrackedChanges();
        trackedChanges.load("items");
        await context.sync();

        trackedChanges.items.forEach((change) => change.reject());
        await context.sync();
    });
}

async function disableAllChanges() {
    await Word.run(async (context) => {
        const revisions = context.document.body.getTrackedChanges();
        revisions.load("items");
        await context.sync();

        for (const revision of revisions.items) {
            revision.reject();
        }
        context.document.trackRevisions = false;
        await context.sync();
    });
}

async function extractDocumentAsJSONAndInsert() {
    return await Word.run(async (context) => {
        try {
            let documentStructure = {
                header: "",
                footer: "",
                document: []
            };

            const sections = context.document.sections;
            sections.load("items");
            await context.sync();

            if (sections.items.length > 0) {
                const header = sections.items[0].getHeader("Primary");
                const footer = sections.items[0].getFooter("Primary");
                header.load("text");
                footer.load("text");
                await context.sync();

                documentStructure.header = header.text ? header.text.trim() : "";
                documentStructure.footer = footer.text ? footer.text.trim() : "";
            }

            // Load paragraphs and tables
            const body = context.document.body;
            const paragraphs = body.paragraphs;
            const tables = body.tables;

            paragraphs.load("items/style,text");
            tables.load("items");
            await context.sync();

            let currentSection = null;
            let index = 1;

            // Process paragraphs
            for (let para of paragraphs.items) {
                let paraText = para.text ? para.text.trim() : "";

                if (!paraText) continue; // Skip empty paragraphs

                // Check if the paragraph is a heading
                if (para.style && para.style.name && para.style.name.startsWith("Heading")) {
                    if (currentSection) {
                        documentStructure.document.push(currentSection);
                    }
                    currentSection = {
                        type: "section",
                        title: paraText,
                        index: index,
                        paragraphs: []
                    };
                } else if (currentSection) {
                    currentSection.paragraphs.push({
                        type: "paragraph",
                        text: paraText,
                        index: index
                    });
                } else {
                    documentStructure.document.push({
                        type: "paragraph",
                        text: paraText,
                        index: index
                    });
                }

                index++;
            }

            if (currentSection) {
                documentStructure.document.push(currentSection);
            }

            // Process tables
            await processTableData(tables, documentStructure, index, context);

                   

            // **Insert JSON into the document**
            const jsonString = JSON.stringify(documentStructure, null, 2);
            // context.document.body.insertParagraph("Extracted JSON:", Word.InsertLocation.end);
            // context.document.body.insertParagraph(jsonString, Word.InsertLocation.end);

            await context.sync();

            console.log("[DEBUG] Extracted Document JSON inserted into the document.");
            return documentStructure;

        } catch (error) {
            console.error("[DEBUG] Error extracting document:", error);
            return { error: error.message };
        }
    });
}

async function processTableData(tables, documentStructure, index, context) {
    for (let tableIndex = 0; tableIndex < tables.items.length; tableIndex++) {
        let table = tables.items[tableIndex];
        table.rows.load("items"); // Load all rows in the table
        await context.sync();

        let tableData = [];

        for (let rowIndex = 0; rowIndex < table.rows.items.length; rowIndex++) {
            let row = table.rows.items[rowIndex];
            row.cells.load("items"); // Load all cells in the row
            await context.sync();

            let rowData = [];

            for (let cellIndex = 0; cellIndex < row.cells.items.length; cellIndex++) {
                let cell = row.cells.items[cellIndex];
                cell.load("text"); // Explicitly load the text property
                await context.sync();

                let cellText = cell.text ? cell.text.trim() : "[Empty]";
                rowData.push(cellText);
            }

            tableData.push(rowData);
        }

        documentStructure.document.push({
            type: "table",
            table: tableData,
            index: index
        });

        index++;
    }
}


function displayExtractedJSON(jsonData) { 
    const jsonOutputElement = document.getElementById("jsonOutput");

    if (jsonOutputElement) {
        jsonOutputElement.textContent = JSON.stringify(jsonData, null, 2);
    } else {
        console.warn("Could not find #jsonOutput element to display JSON.");
    }
}



async function applyReplace(change) {
    await Word.run(async (context) => {
        try {
            // Step 1: Enable track changes before making modifications
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();

            // Step 2: Load all paragraphs
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");
            await context.sync();

            // Step 3: Validate paragraph index
            const paragraphIndex = change.paragraph_index - 1; // Convert 1-based index to 0-based
            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
                await insertDebugMessage(`Error: Paragraph index ${change.paragraph_index} out of range.`);
                return;
            }

            const targetParagraph = paragraphs.items[paragraphIndex];
            targetParagraph.load("text");
            await context.sync();

            const paragraphText = targetParagraph.text;

            // Step 4: Normalize the texts for comparison
            const normalizedParagraphText = paragraphText.trim().replace(/\r?\n|\r/g, "").replace(/\s+/g, " ");
            const normalizedOriginalText = change.original_text.trim().replace(/\r?\n|\r/g, "").replace(/\s+/g, " ");

            // Debugging: Log normalized values
            // await insertDebugMessage(
            //     `Normalized paragraphText: '${normalizedParagraphText}', normalizedOriginalText: '${normalizedOriginalText}'`
            // );
            const trimmedUpdatedText = change.updated_text.trim().replace(/\r?\n|\r/g, "");

            // Step 5: Check for the presence of original text
            if (!normalizedParagraphText.includes(normalizedOriginalText)) {
                await insertDebugMessage(
                    `Error: Normalized original text '${normalizedOriginalText}' not found in paragraph ${change.paragraph_index}.`
                );
                return;
            }

            // Step 6: Perform the replacement using raw paragraphText
            const updatedParagraphText = paragraphText.replace(normalizedOriginalText, trimmedUpdatedText);

            // Debugging: Log the updated text
            // await insertDebugMessage(`Updated Paragraph Text: '${updatedParagraphText}'`);

            // Step 7: Update the paragraph with the replaced text
            targetParagraph.clear(); // Clear the current paragraph content
            targetParagraph.insertText(updatedParagraphText, Word.InsertLocation.replace); // Replace with updated text
            await context.sync();

            // Step 8: Log success
            // await insertDebugMessage(
            //     `Applied change_id ${change.change_id}: Replaced '${change.original_text}' with '${change.updated_text}' in paragraph ${change.paragraph_index}.`
            // );
        } catch (error) {
            // Step 9: Log any errors
            await insertDebugMessage(`Error applying replace for change_id ${change.change_id}: ${error.message}`);
        }
    });
}


async function applyAddOrUpdate(change) {

    // await insertDebugMessage(`reacged the apply or update method `);

    await Word.run(async (context) => {
        try {
            // Step 1: Log the change object
            // console.log("Change object received:", change);

            // Step 2: Enable track changes
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();

            // Step 3: Load paragraphs
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");
            await context.sync();

            // await insertDebugMessage("Total paragraphs in the document:", paragraphs.items.length);

            // Step 4: Validate text and determine action
            const textToApply = change.inserted_text || change.updated_text || change.content;
            if (!textToApply) {
                throw new Error(`No valid text provided for change_id ${change.change_id}.`);
            }

            // console.log("Text to insert or update:", textToApply);

            if (change.action === "add") {
                // Step 5: Handle adding content
                if (change.paragraph_index && change.paragraph_index - 1 < paragraphs.items.length) {
                    const targetParagraph = paragraphs.items[change.paragraph_index - 1];
                    const position = change.position || "after";

                    let insertLocation;
                    switch (position) {
                        case "before":
                            insertLocation = Word.InsertLocation.before;
                            break;
                        case "after":
                            insertLocation = Word.InsertLocation.after;
                            break;
                        case "start":
                            insertLocation = Word.InsertLocation.start;
                            break;
                        case "end":
                            insertLocation = Word.InsertLocation.end;
                            break;
                        default:
                            console.warn("Invalid position. Defaulting to append at the end.");
                            insertLocation = Word.InsertLocation.end;
                    }

                    console.log(`Inserting text at position: ${position}`);
                    targetParagraph.insertText(textToApply, insertLocation);
                } else {
                    // Append at the end if paragraph index is invalid
                    console.warn("Invalid paragraph index. Appending text at the end.");
                    context.document.body.insertParagraph(textToApply, Word.InsertLocation.end);
                }
            } else if (change.action === "update") {
                // Step 6: Handle updating content
                if (change.paragraph_index && change.paragraph_index - 1 < paragraphs.items.length) {
                    const targetParagraph = paragraphs.items[change.paragraph_index - 1];
                    targetParagraph.load("text");
                    await context.sync();

                    // Replace the paragraph content
                    targetParagraph.clear();
                    targetParagraph.insertText(textToApply, Word.InsertLocation.replace);
                } else {
                    console.warn("Invalid paragraph index for update. Skipping update.");
                }
            } else {
                throw new Error(`Unsupported action '${change.action}' for change_id ${change.change_id}.`);
            }

            await context.sync();

            // Log success
            console.log(`Successfully applied ${change.action} for change_id ${change.change_id}.`);
        } catch (error) {
            console.error(`Error applying ${change.action} for change_id ${change.change_id}: ${error.message}`);
        }
    });
}

async function applyChange(change) {
    if (change.action === "replace") {
        await applyReplace(change);
    } else if (change.action === "add" || change.action === "update") {
        await applyAddOrUpdate(change);
    } else {
        await insertDebugMessage(`Error: Unsupported action '${change.action}' for change_id ${change.change_id}.`);
    }
}


async function sendDocumentJSONToAPI(instruction) {
    await Word.run(async (context) => {
        try {
            // Step 1: Extract document content as JSON
            const documentJSON = await extractDocumentAsJSONAndInsert();

            if (!documentJSON || Object.keys(documentJSON).length === 0) {
                console.error("[DEBUG] No document content extracted or JSON is empty.");
                await insertDebugMessage("Error: Unable to extract document content.");
                return;
            }
            
            const storedDetails = sessionStorage.getItem('companyDetails');
            
            if (!storedDetails) {
                insertDebugMessage("Please login to use this add-in");
                document.getElementById("userInfo").innerHTML = 
                    "Please login to use this feature";
                return; // Stop execution if no company details
            }

            const companyDetails = JSON.parse(storedDetails);
            const company = companyDetails.companyName; // Get company name

            // Step 2: Prepare the payload
            const payload = {
                instruction: instruction,
                document_content: documentJSON, 
                companyName : company// Send raw JSON, not stringified
            };

            // Log the payload for testing
            console.log("[DEBUG] Payload being sent to API:", payload);

            // Step 3: Send to API
            const response = await axios.post(
                "https://a169-2607-fea8-fc01-7009-d074-5ac-b353-5829.ngrok-free.app/process_json",
                payload,
                { headers: { "Content-Type": "application/json" } }
            );

            // Step 4: Handle response
            if (response.status !== 200) {
                throw new Error(`API returned status ${response.status}`);
            } 
            
            // await insertDebugMessage(`Raw API response: ${JSON.stringify(response.data)}`);

             // Access and validate changes
            const changesArray = response.data?.changes; // Access changes from API response

            // Ensure changesArray is an array
            if (!Array.isArray(changesArray)) {
                await insertDebugMessage(`Error: 'changes' is not an array. Found type: ${typeof changesArray}`);
                return;
            }

            
            displayProposedChanges(changesArray);

        } catch (error) {
            console.error("[DEBUG] Error in sendDocumentJSONToAPI:", error);
            await insertDebugMessage(`Error: ${error.message}`);
        }
    });
}


async function displayProposedChanges(changes) {
    const container = document.getElementById("proposedChangesContainer");
    container.style.display = 'block';
    container.innerHTML = ""; // Clear previous content

    if (!Array.isArray(changes) || changes.length === 0) {
        container.innerHTML = "<p>No changes detected.</p>";
        await insertDebugMessage("No changes detected by the API.");
        return;
    }

    changes.forEach((change) => {
        const changeItem = document.createElement("div");
        changeItem.className = "change-item";

        const text = document.createElement("p");

        if (change.action === "replace") {
            if (!change.original_text || !change.updated_text) {
                insertDebugMessage(`Error: Missing data for replace change_id ${change.change_id}.`);
                return;
            }

            text.textContent = `Paragraph ${change.paragraph_index}: ${change.original_text}`;

            const proposed = document.createElement("p");
            proposed.innerHTML = `<strong>Proposed Change:</strong> ${change.updated_text}`;

            changeItem.appendChild(text);
            changeItem.appendChild(proposed);
        } 
        else if (change.action === "add" || change.action === "update") {
            if (!change.inserted_text && !change.updated_text) {
                insertDebugMessage(`Error: Missing inserted_text or updated_text for change_id ${change.change_id}.`);
                return;
            }

            text.textContent = `Paragraph ${change.paragraph_index || "End"}:`;

            const proposed = document.createElement("p");
            proposed.innerHTML = `<strong>Inserted/Updated Text:</strong> ${change.inserted_text || change.updated_text}`;

            changeItem.appendChild(text);
            changeItem.appendChild(proposed);
        } 
        else {
            insertDebugMessage(`Error: Unknown action monisha 1 '${change.action}' for change_id ${change.change_id}.`);
            return;
        }

        // insertDebugMessage(`going before `);

        // Create Accept/Reject buttons
        const acceptButton = document.createElement("button");
        acceptButton.textContent = "Accept";
        acceptButton.onclick = () => applyChange(change);
        

        const rejectButton = document.createElement("button");
        rejectButton.textContent = "Reject";
        rejectButton.onclick = () => insertDebugMessage(`Rejected change_id ${change.change_id}`);


        changeItem.appendChild(acceptButton);
        changeItem.appendChild(rejectButton);
        container.appendChild(changeItem);
    });

    // await insertDebugMessage(`Preparing to display ${changes.length} proposed changes.`);
}



async function generateAIChangesWithContext() {
    const userPrompt = document.getElementById("aiPromptInput").value;
    if (!userPrompt) {
        console.error("No prompt provided.");
        // await insertDebugMessage("No prompt provided for AI generation.");
        return;
    }
    insertDebugMessage("does it reach here first of all ")
    // await insertDebugMessage(`Button clicked. Starting AI generation with prompt: ${userPrompt}`);

    // Call the function to send the document content and process it
    // await sendDocumentContentToAPI(userPrompt);
    await sendDocumentJSONToAPI(userPrompt);
}


async function insertDebugMessage(message) {
    await Word.run(async (context) => {
        const body = context.document.body;
        body.insertText(`[DEBUG]: ${message}\n`, Word.InsertLocation.start);
        await context.sync();
        console.log(`[DEBUG]: ${message}`); // Also log to the browser console
    });
}


function arrayBufferToBase64(buffer) {
    let binary = "";
    let bytes = new Uint8Array(buffer);
    let len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}
