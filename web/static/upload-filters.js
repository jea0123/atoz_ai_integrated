const UploadFilters = (() => {
  const DEFAULT_IGNORED_FOLDER_KEYS = new Set(["bak", "backup", "old"]);
  const ARTIFACT_EXTENSIONS = new Set([".hwp", ".hwpx", ".docx", ".docm", ".pptx", ".pptm", ".xlsx", ".xlsm", ".xltx", ".xltm"]);
  const CHECK_EXTENSIONS = new Set([...ARTIFACT_EXTENSIONS, ".doc", ".xls", ".ppt", ".potx", ".potm", ".ppsx", ".ppsm"]);

  function normalizePathPart(value) {
    return String(value || "")
      .normalize("NFKC")
      .replace(/\s+/g, "")
      .toLowerCase();
  }

  function relativePath(file) {
    return file.webkitRelativePath || file.name;
  }

  function pathParts(file) {
    return relativePath(file)
      .replaceAll("\\", "/")
      .split("/")
      .filter(Boolean);
  }

  function extension(file) {
    const name = file.name || "";
    const dotIndex = name.lastIndexOf(".");
    return dotIndex >= 0 ? name.slice(dotIndex).toLowerCase() : "";
  }

  function hasIgnoredFolder(file, ignoredFolderKeys) {
    return pathParts(file)
      .slice(0, -1)
      .map(normalizePathPart)
      .some((part) => ignoredFolderKeys.has(part));
  }

  function selectedFiles(fileList, supportedExtensions, ignoredFolderKeys = DEFAULT_IGNORED_FOLDER_KEYS) {
    return Array.from(fileList || [])
      .filter((file) => supportedExtensions.has(extension(file)))
      .filter((file) => !hasIgnoredFolder(file, ignoredFolderKeys));
  }

  return {
    ARTIFACT_EXTENSIONS,
    CHECK_EXTENSIONS,
    relativePath,
    selectedFiles,
  };
})();
