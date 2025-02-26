/* global document, Office, Word */
import axios from 'axios';


let currentPlaceholders = {};

// Event listeners for buttons in the task pane
document.getElementById("loadTemplate").addEventListener("click", loadTemplate);
document.getElementById("applyEditedPlaceholders").addEventListener("click", updatePlaceholders);
document.getElementById("enableTrackChanges").addEventListener("click", enableTrackChanges);
document.getElementById("applyRedlines").addEventListener("click", applyRedlines);
document.getElementById("generateAIChanges").addEventListener("click", generateAIChanges);
// document.getElementById("generateAIChangesWithContext").addEventListener("click", generateAIChangesWithContext);

document.getElementById("disableAllChanges").addEventListener("click", disableAllChanges);
document.getElementById("listTrackedChanges").addEventListener("click", listTrackedChanges);
document.getElementById("acceptSelectedChanges").addEventListener("click", acceptSelectedTrackedChanges);
document.getElementById("rejectSelectedChanges").addEventListener("click", rejectSelectedTrackedChanges);
document.getElementById("acceptAllChanges").addEventListener("click", acceptAllTrackedChanges);
document.getElementById("rejectAllChanges").addEventListener("click", rejectAllTrackedChanges);
document.getElementById("templatesBtn").addEventListener("click", getTemplatesAndPopulateDropdown);
document.getElementById("openTemplateBtn").addEventListener("click", fetchAndOpenTemplate);
document.getElementById("loginButton").addEventListener("click", initializeTaskPane);


// // document.getElementById("saveAdminSettings").addEventListener("click", saveAdminSettings);


Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Attach event listeners
        document.getElementById("docxFileInput").addEventListener("change", handleDocxUpload);
        document.getElementById("applyEditedPlaceholders").addEventListener("click", mergeAndInsertTemplate);

    }
});

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


// --------------------- now we are working on the login functionality ---------------------------


// Ensure Office is ready before running your code.
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeTaskPane();
  }
});

// Main initialization: get user email and toggle admin UI.
async function initializeTaskPane() {
  try {
    const userEmail = await getUserEmail();
    insertDebugMessage("Logged-in user email: " + userEmail);
    console.log("Logged-in user email:", userEmail);

    const isAdminUser = isAdmin(userEmail);
    insertDebugMessage("Is admin: " + isAdminUser);
    toggleAdminForm(isAdminUser);
  } catch (error) {
    console.error("Error initializing task pane:", error);
    insertDebugMessage(`Error initializing task pane: " + ${error.message}`);
  }
}
Office.initialize = function (reason) {
    $(function () { 
        insertDebugMessage("prii em al   ", Office.context.mailbox.userProfile.emailAddress);
    }
}
// Retrieves the signed-in user's email using Office SSO and Microsoft Graph.
function getUserEmail() {
        insertDebugMessage("inside get userPrompt: " );

  return new Promise((resolve, reject) => {
    Office.context.auth.getAccessTokenAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const token = result.value;
            insertDebugMessage(`do we have the token ${token}`);

        // Call Microsoft Graph to get the user's profile.
        fetch("https://graph.microsoft.com/v1.0/me", {
          headers: {
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json"
          }
        })
          .then((response) => {
            if (!response.ok) {
              reject(new Error("Failed to fetch user profile: " + response.statusText));
            }
            return response.json();
          })
          .then((data) => {
            const email = data.mail || data.userPrincipalName;
            if (email) {
              resolve(email);
            } else {
              reject(new Error("No email found in user profile."));
            }
          })
          .catch((error) => reject(error));
      } else {
        reject(new Error("Failed to get token: " + JSON.stringify(result.error)));
      }
    });
  });
}

// List of admin emails. Adjust these values as needed.
function isAdmin(email) {
  const adminEmails = [
    "c0914148@mylambton.ca",
    "admin2@example.com",
    "monisha@kubetools.onmicrosoft.com"
  ];
  return adminEmails.includes(email.toLowerCase());
}

// Toggle visibility of the admin settings form based on admin status.
function toggleAdminForm(isAdminUser) {
  const adminPage = document.getElementById("adminPage");
  const loginButton = document.getElementById("loginButton");

  if (isAdminUser) {
    adminPage.style.display = "block"; // Show admin settings.
    loginButton.style.display = "none"; // Hide login button.
  } else {
    adminPage.style.display = "none";  // Hide admin settings.
    loginButton.style.display = "block"; // Show login button.
  }
}

// Debug helper: log messages both to console and (if exists) a debug div.

// Main function: Called when Office.js is ready

// List of admin emails (replace with your actual admin emails or fetch from a database)

// Function to get an access token using SSO
// Function to get an access token using SSO



// List of admin emails (replace with your actual admin emails or fetch from a database)


// const msalInstance = new PublicClientApplication(msalConfig);

// async function getUserToken() {
//   const request = { scopes: ["User.Read"] };
//   try {
//     const response = await msalInstance.acquireTokenSilent(request);
//     return response.accessToken;
//   } catch (error) {
//     if (error instanceof InteractionRequiredAuthError) {
//       const response = await msalInstance.acquireTokenPopup(request);
//       return response.accessToken;
//     } else {
//       throw error;
//     }
//   }
// }

// async function getUserEmail() {
//     insertDebugMessage(`Fhere is user email function `);

//     try {
//         // Get the access token
//         const accessToken = await getAccessToken();
//         // const accessToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6InA5bnc2NGJodlJPRDVISWlzRTFJU1VJWTZoVW9HYVNHTlBiTEZ1LU16MmciLCJhbGciOiJSUzI1NiIsIng1dCI6ImltaTBZMnowZFlLeEJ0dEFxS19UdDVoWUJUayIsImtpZCI6ImltaTBZMnowZFlLeEJ0dEFxS19UdDVoWUJUayJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8yMjJiNGZmMy0xYjNjLTQwNTEtYjJiOC03NjM0OWVlMzc4OGMvIiwiaWF0IjoxNzQwNTI4MTk1LCJuYmYiOjE3NDA1MjgxOTUsImV4cCI6MTc0MDUzMjA5NSwiYWlvIjoiazJSZ1lGRExmN1hQa3QzWHRPUGI1UGdVOGZnN0FBPT0iLCJhcHBfZGlzcGxheW5hbWUiOiJXb3JkIFBsdWdpbiBBUEkiLCJhcHBpZCI6IjJhYzcyODllLTE5ZWMtNDgzMi1iZmZmLTE2YzZhOGI0ZThiMiIsImFwcGlkYWNyIjoiMSIsImlkcCI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzIyMmI0ZmYzLTFiM2MtNDA1MS1iMmI4LTc2MzQ5ZWUzNzg4Yy8iLCJpZHR5cCI6ImFwcCIsIm9pZCI6ImQwODdkYTYzLWUxOGEtNDEzZS05YzhhLTIwOGU4OWI5ZGIzNiIsInJoIjoiMS5BVzhCODA4cklqd2JVVUN5dUhZMG51TjRqQU1BQUFBQUFBQUF3QUFBQUFBQUFBQndBUUJ2QVEuIiwicm9sZXMiOlsiVXNlci5SZWFkQmFzaWMuQWxsIiwiRmlsZXMuUmVhZFdyaXRlLkFwcEZvbGRlciIsIlVzZXIuUmVhZFdyaXRlLkFsbCIsIkZpbGVzLlNlbGVjdGVkT3BlcmF0aW9ucy5TZWxlY3RlZCIsIlNpdGVzLlJlYWQuQWxsIiwiU2l0ZXMuUmVhZFdyaXRlLkFsbCIsIkZpbGVzLlJlYWRXcml0ZS5BbGwiLCJVc2VyLlJlYWQuQWxsIiwiRmlsZXMuUmVhZC5BbGwiLCJFeHRlcm5hbFVzZXJQcm9maWxlLlJlYWQuQWxsIiwiRXh0ZXJuYWxVc2VyUHJvZmlsZS5SZWFkV3JpdGUuQWxsIl0sInN1YiI6ImQwODdkYTYzLWUxOGEtNDEzZS05YzhhLTIwOGU4OWI5ZGIzNiIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6IjIyMmI0ZmYzLTFiM2MtNDA1MS1iMmI4LTc2MzQ5ZWUzNzg4YyIsInV0aSI6IlJaWGl4dks1aGthRkt0XzFOVEdCQVEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjA5OTdhMWQwLTBkMWQtNGFjYi1iNDA4LWQ1Y2E3MzEyMWU5MCJdLCJ4bXNfaWRyZWwiOiI0IDciLCJ4bXNfdGNkdCI6MTczNzA3MDE3NX0.ml6xtpa-oPTay9RexC_poZeAj4J4efFZU4UbCsqdvK5Maf_KwTOuaFCy4B8IiPURtWWHSjISFBgXm7nS46_lt18PODcsuUgkrSbVb39UHjsDcTyNHF1sHKMV9LtLwnlPRBBXNO-lJNUv8unatdw2XI0W5BUkWWDyY_SHqfe4U9_R0q5jsTkGjJUdpDGwDl1rltrMb_MqbJm_YqPoSo9Lfshz0x-_1dsWg9b8_38w9st8_eV5PQ4GFzE5hmdfjsFhkBDAdZ8aFqUaXBxp2qrQEEgEmFuuXkDh120hApdjyBinRz7Fn5160C6Wpp4mlN4owsGsqZo1zF9NgXbg434uGQ"
//         insertDebugMessage(`Access token:", ${accessToken}`); // Log the access token

//         // Call Microsoft Graph API to fetch the user's profile
//         const response = await fetch("https://graph.microsoft.com/v1.0/me", {
//             headers: {
//                 Authorization: `Bearer ${accessToken}`,
//             },
//         });

//         if (!response.ok) {
//             const error = new Error("Failed to fetch user profile.");
//             console.error("Failed to fetch user profile:", response.statusText);
//             throw error;
//         }

//         const userProfile = await response.json();
//         insertDebugMessage(`User profile:", ${userProfile}`); // Log the user profile
//         return userProfile.mail || userProfile.userPrincipalName; // Return the user's email
//     } catch (error) {
//         console.error("Error fetching user email:", error);
//         throw error;
//     }
// }


// function handleError(error) {
//     console.error("Error details:", error); // Log the full error object
//     console.error("Error message:", error.message); // Log the error message
//     console.error("Error stack:", error.stack); // Log the error stack trace
//     insertDebugMessage("Error: " + error.message); // Display the error in the UI
// }


// Office.onReady((info) => {
//     if (info.host === Office.HostType.Word) {
//         insertDebugMessage("Office.js is ready. Attempting to get SSO token...");
//         getSSOTokenAndFetchUser();
//     }
// });

// // Function to get SSO token using Office.context.auth.getAccessTokenAsync
// function getSSOTokenAndFetchUser() {
//     Office.context.auth.getAccessTokenAsync((result) => {
//         if (result.status === Office.AsyncResultStatus.Succeeded) {
//             let token = result.value;
//             insertDebugMessage("Access token retrieved successfully.");
//             // Now use the token to call Microsoft Graph API
//             fetchUserInfo(token);
//         } else {
//             insertDebugMessage("Failed to get access token: " + JSON.stringify(result.error));
//         }
//     });
// }

// // Function to call Microsoft Graph API and fetch the user's email
// function fetchUserInfo(accessToken) {
//     insertDebugMessage("Fetching user info from Microsoft Graph...");
//     fetch("https://graph.microsoft.com/v1.0/me", {
//         headers: { Authorization: `Bearer ${accessToken}` }
//     })
//     .then(response => response.json())
//     .then(data => {
//         // Log the full response for debugging
//         insertDebugMessage("Microsoft Graph API response: " + JSON.stringify(data));
//         let userEmail = data.mail || data.userPrincipalName;
//         if (!userEmail) {
//             insertDebugMessage("No email found in the Microsoft Graph response.");
//             return;
//         }
//         insertDebugMessage("User email retrieved: " + userEmail);
//         // At this point, you have the user's email.
//         // For example, extract the domain:
//         let domain = userEmail.split("@")[1];
//         insertDebugMessage("User domain: " + domain);
//         // Next, you can use the email (or domain) to look up company-specific settings from your backend.
//     })
//     .catch(error => {
//         insertDebugMessage("Error fetching user info from Graph: " + error);
//     });
// }


const templates = {
    template1: {
        header: "Header: {{Name}} - {{Date}}",
        body: "Dear {{Name}},\nToday is {{Date}}.\nBest regards,\nYour Company.",
        footer: "Footer: Sincerely, {{Signature}}",
        placeholders: ["Name", "Date", "Signature"],
    },
    template2: {
        header: "Appointment Reminder: {{Name}} - {{Date}}",
        body: "Hello {{Name}},\nYour appointment is scheduled for {{Date}}.\nThank you!",
        footer: "Footer: Best regards,\n{{Signature}}",
        placeholders: ["Name", "Date", "AppointmentTime", "Signature"],
    },
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        // Auto-fetch templates as soon as the task pane is ready
        getTemplatesAndPopulateDropdown();
    }
});


async function loadTemplate() {
    const templateKey = document.getElementById("templateSelector").value;
    const template = templates[templateKey];

    await Word.run(async (context) => {
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();

        sections.items.forEach((section) => {
            const header = section.getHeader("Primary");
            header.clear();
            header.insertText(template.header, Word.InsertLocation.replace);

            const body = section.body;
            body.clear();
            body.insertText(template.body, Word.InsertLocation.start);

            const footer = section.getFooter("Primary");
            footer.clear();
            footer.insertText(template.footer, Word.InsertLocation.replace);
        });

        await context.sync();
    });

    generateEditFields(template.placeholders);
}

async function getTemplatesAndPopulateDropdown() {
    try {
      // Step 1: Fetch data from the backend endpoint with Axios
    //   const response = await axios.get("https://bca5-142-126-182-191.ngrok-free.app/api/templates");

    const response = await axios.get("https://bca5-142-126-182-191.ngrok-free.app/api/templates", {
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

        console.log(`Fetching template with ID: ${selectedTemplateId}`);

        // Step 1: Fetch the .docx file from backend
        const response = await axios.get(`https://bca5-142-126-182-191.ngrok-free.app/api/templates/${selectedTemplateId}`, {
            headers: { "ngrok-skip-browser-warning": "true" },
            responseType: "arraybuffer" // ⚠️ Change response type to arraybuffer
        });

        console.log("Template file received from API:", response);

        // Step 2: Convert ArrayBuffer to Base64
        const base64Data = arrayBufferToBase64(response.data);

        console.log("Converted file to Base64, attempting to insert into Word...");

        // Step 3: Insert into Word
        await Word.run(async (context) => {
            const body = context.document.body;
            body.clear();
            body.insertFileFromBase64(base64Data, Word.InsertLocation.start);
            await context.sync();
        });

        console.log("✅ Template successfully inserted into Word!");
        generateEditFields(template.placeholders);

    } catch (error) {
        console.error("❌ Error fetching template:", error);
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


async function generateAIChanges() {
    const userPrompt = document.getElementById("aiPromptInput").value;
    if (!userPrompt) {
        console.error("No prompt provided.");
        return;
    }

    await Word.run(async (context) => {
        try {
            const aiResponse = await getAIResponse(userPrompt);
            if (!aiResponse) {
                console.error("AI response is null or undefined.");
                return;
            }

            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();

            if (!selection.text) {
                console.warn("No text selected. Inserting at cursor position.");
            }

            selection.insertText(aiResponse, Word.InsertLocation.replace);
            await context.sync();

            console.log("AI suggestion inserted successfully.");
        } catch (error) {
            console.error("Error in generateAIChanges:", error);
        }
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


// Attach this function to a button click for testing
document.getElementById("testWelcomeButton").addEventListener("click", callAPI);

async function callAPI() {
    try {
        const response = await axios.post(
            "https://0eac-2607-fea8-580-c300-dddd-7660-2cd1-9b7c.ngrok-free.app/process_instruction",
            {
                instruction: "Replace Monisha with Mon",
                document_content: "Sample document content",
            },
            {
                headers: {
                    "Content-Type": "application/json",
                },
            }
        );
        // await insertDebugMessage("API Response:", response.data);
    } catch (error) {
        // await insertDebugMessage("API Call Error:", error.message);
    }
}

async function sendDocumentContentToAPI(instruction) {
    await Word.run(async (context) => {
        try {
            
            // Step 1: Retrieve document content
            // const documentJSON = await extractDocumentAsJSONAndInsert();

                    // ✅ Show JSON before sending to API
        // displayExtractedJSON(documentJSON);



            const body = context.document.body;
            body.load("text");
            await context.sync();

            const documentContent = body.text;
            // await insertDebugMessage(`Document content before API call: ${documentContent}`);

            // Step 2: Prepare payload and send API request
            const payload = {
                instruction: instruction,
                document_content: btoa(documentContent), // Encode as Base64
            };

            // await insertDebugMessage(`Payload sent to API: ${JSON.stringify(payload)}`);

            const response = await axios.post(
                "https://9ff9-142-189-167-212.ngrok-free.app/process_instruction",
                payload,
                { headers: { "Content-Type": "application/json" } }
            );

            if (response.status !== 200) {
                throw new Error(`API returned status ${response.status}`);
            }

            // Step 3: Debug the full API response
            // await insertDebugMessage(`Raw API response: ${JSON.stringify(response.data)}`);

            // **Correct extraction of changes and updated_document**
            const updated_document_obj = response.data?.updated_document || {};
            const changes = updated_document_obj.changes || [];
            const updated_document = updated_document_obj.updated_document || "";

            // await insertDebugMessage(`Parsed changes: ${JSON.stringify(changes)}`);
            // await insertDebugMessage(`Parsed updated document type: ${typeof updated_document}`);

            // Validate API response
            if (!Array.isArray(changes)) {
                await insertDebugMessage("Error: API did not return an array for 'changes'.");
                return;
            }

            if (typeof updated_document !== "string") {
                await insertDebugMessage("Error: API did not return Base64 string for 'updated_document'.");
                return;
            }

            // Step 4: Display proposed changes for review
            displayProposedChanges(changes);

        } catch (error) {
            await insertDebugMessage(`Error: ${error.message}`);
        }
    });
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

            // Log the JSON for debugging
            console.log("[DEBUG] Document JSON extracted:", documentJSON);

            // Step 2: Prepare the payload
            const payload = {
                instruction: instruction,
                document_content: documentJSON, // Send raw JSON, not stringified
            };

            // Log the payload for testing
            console.log("[DEBUG] Payload being sent to API:", payload);

            // Step 3: Send to API
            const response = await axios.post(
                "https://bca5-142-126-182-191.ngrok-free.app/process_json",
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



document.getElementById("generateAIChangesWithContext").addEventListener("click", async () => {
    const userPrompt = document.getElementById("aiPromptInput").value;
    if (!userPrompt) {
        console.error("No prompt provided.");
        // await insertDebugMessage("No prompt provided for AI generation.");
        return;
    }

    // await insertDebugMessage(`Button clicked. Starting AI generation with prompt: ${userPrompt}`);

    // Call the function to send the document content and process it
    // await sendDocumentContentToAPI(userPrompt);
    await sendDocumentJSONToAPI(userPrompt);
});




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
