/** crmDriveFiles.gs */

function crm_uploadFileToDriveAndRegister(payload) {
  // payload:
  // {
  //   matterId,
  //   title,
  //   type,
  //   status,
  //   notes,
  //   fileName,
  //   mimeType,
  //   base64Data
  // }

  if (!payload) throw new Error("crm_uploadFileToDriveAndRegister: missing payload");
  if (!payload.matterId) throw new Error("crm_uploadFileToDriveAndRegister: missing matterId");
  if (!payload.fileName) throw new Error("crm_uploadFileToDriveAndRegister: missing fileName");
  if (!payload.base64Data) throw new Error("crm_uploadFileToDriveAndRegister: missing base64Data");

  const matter = crm_getMatterById(payload.matterId);
  if (!matter) throw new Error("crm_uploadFileToDriveAndRegister: matter not found");

  const clientId = String(matter.CLIENT_ID || "");
  if (!clientId) throw new Error("crm_uploadFileToDriveAndRegister: matter has no clientId");

  const bytes = Utilities.base64Decode(payload.base64Data);
  const mimeType = payload.mimeType || "application/octet-stream";
  const blob = Utilities.newBlob(bytes, mimeType, payload.fileName);

  // Try to upload to matter folder if available
  let file = null;
  if (matter.FOLDER_URL) {
    const upFolderId = extractFolderIdFromUrl_(matter.FOLDER_URL);
    if (upFolderId) {
      try {
        const uploadsFolder = crm_ensureDriveFolder_(upFolderId, "Uploads");
        if (uploadsFolder) {
          const uploadFolder = DriveApp.getFolderById(uploadsFolder.folderId);
          file = uploadFolder.createFile(blob.setName(payload.fileName));
        }
      } catch (e) {
        logInfo_("UPLOAD_FOLDER_ERROR", "Failed to upload to matter folder, falling back to root", {
          matterId: payload.matterId,
          error: e.message,
        });
      }
    }
  }

  // Fallback to root folder if matter folder unavailable
  if (!file) {
    file = crm_saveBlobToDrive_(blob, payload.fileName);
  }

  const fileId = file.getId();
  const fileUrl = file.getUrl();
  const title = (payload.title || payload.fileName || "Uploaded file").trim();

  const isGoogleDoc = mimeType.indexOf("application/vnd.google-apps.document") !== -1 ||
                      fileUrl.indexOf("/document/") !== -1;

  const isPdf = mimeType === "application/pdf" || /\.pdf$/i.test(payload.fileName);

  const docRes = crm_addDocument({
    clientId: clientId,
    matterId: payload.matterId,
    type: payload.type || "GENERAL",
    status: payload.status || "READY",
    title: title,
    docUrl: isGoogleDoc ? fileUrl : "",
    pdfUrl: isPdf ? fileUrl : "",
    fileId: fileId,
    createdBy: getActiveUserEmail_() || "unknown",
    notes: payload.notes || "",
  });

  crm_logActivity({
    action: "DOCUMENT_FILE_UPLOADED",
    message: `Document file uploaded: ${payload.fileName}`,
    clientId: clientId,
    matterId: payload.matterId,
    meta: {
      fileId: fileId,
      fileUrl: fileUrl,
      fileName: payload.fileName,
      mimeType: mimeType,
      docId: docRes.docId || "",
    },
  });

  return {
    ok: true,
    fileId: fileId,
    fileUrl: fileUrl,
    docId: docRes.docId || "",
    matterId: payload.matterId,
    clientId: clientId,
  };
}

function crm_saveBlobToDrive_(blob, fileName) {
  const c = cfg_();
  const folderId = c.DRIVE && c.DRIVE.ROOT_FOLDER_ID ? String(c.DRIVE.ROOT_FOLDER_ID).trim() : "";

  if (folderId) {
    const folder = DriveApp.getFolderById(folderId);
    return folder.createFile(blob.setName(fileName));
  }

  return DriveApp.createFile(blob.setName(fileName));
}

/**
 * Get or create a folder by name under parent folder.
 * Safe: only creates if not found. Returns folder object & URL.
 */
function crm_ensureDriveFolder_(parentFolderId, folderName) {
  if (!parentFolderId || !folderName) return null;

  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const folders = parentFolder.getFoldersByName(folderName);

    if (folders.hasNext()) {
      const existing = folders.next();
      return {
        folderId: existing.getId(),
        folderUrl: existing.getUrl(),
        folder: existing,
      };
    }

    const newFolder = parentFolder.createFolder(folderName);
    return {
      folderId: newFolder.getId(),
      folderUrl: newFolder.getUrl(),
      folder: newFolder,
    };
  } catch (e) {
    logInfo_("DRIVE_FOLDER_ERROR", "Failed to ensure folder: " + folderName, {
      parentFolderId,
      folderName,
      error: e.message,
    });
    return null;
  }
}

/**
 * Get or create a client folder: Clients/{CLIENT_ID__FullName}
 * Safe: only creates if not already in Drive. Returns folder URL if successful.
 */
function crm_getOrCreateClientFolder(clientId, clientFullName) {
  if (!clientId || !clientFullName) return null;

  const c = cfg_();
  const rootFolderId = c.DRIVE && c.DRIVE.ROOT_FOLDER_ID ? String(c.DRIVE.ROOT_FOLDER_ID).trim() : "";

  if (!rootFolderId) {
    logInfo_("DRIVE_NO_ROOT", "No root folder configured for client folder creation", { clientId });
    return null;
  }

  try {
    const rootFolder = DriveApp.getFolderById(rootFolderId);

    const clientsFolderRes = crm_ensureDriveFolder_(rootFolderId, "Clients");
    if (!clientsFolderRes) return null;

    const clientFolderName = clientId + "__" + clientFullName.substring(0, 40).trim();
    const clientFolderRes = crm_ensureDriveFolder_(clientsFolderRes.folderId, clientFolderName);

    if (!clientFolderRes) return null;

    crm_logActivity({
      action: "CLIENT_FOLDER_CREATED",
      message: `Client folder created: ${clientFolderName}`,
      clientId: clientId,
      meta: {
        folderId: clientFolderRes.folderId,
        folderUrl: clientFolderRes.folderUrl,
      },
    });

    return clientFolderRes.folderUrl;
  } catch (e) {
    logInfo_("DRIVE_CLIENT_FOLDER_ERROR", "Failed to create client folder", {
      clientId,
      clientFullName,
      error: e.message,
    });
    return null;
  }
}

/**
 * Get or create a matter folder: {clientFolder}/02_Matters/{MATTER_ID__ShortTitle}
 * Safe: clientFolderId must be provided. Returns folder URL if successful.
 */
function crm_getOrCreateMatterFolder(matterId, matterTitle, mattersFolderId) {
  if (!matterId || !matterTitle || !mattersFolderId) return null;

  try {
    const matterFolderName = matterId + "__" + matterTitle.substring(0, 40).trim();
    const matterFolderRes = crm_ensureDriveFolder_(mattersFolderId, matterFolderName);

    if (!matterFolderRes) return null;

    crm_ensureDriveFolder_(matterFolderRes.folderId, "Uploads");

    crm_logActivity({
      action: "MATTER_FOLDER_CREATED",
      message: `Matter folder created: ${matterFolderName}`,
      matterId: matterId,
      meta: {
        folderId: matterFolderRes.folderId,
        folderUrl: matterFolderRes.folderUrl,
      },
    });

    return matterFolderRes.folderUrl;
  } catch (e) {
    logInfo_("DRIVE_MATTER_FOLDER_ERROR", "Failed to create matter folder", {
      matterId,
      matterTitle,
      mattersFolderId,
      error: e.message,
    });
    return null;
  }
}