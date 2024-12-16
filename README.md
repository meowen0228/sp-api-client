# SharePoint Node.js Library

A Node.js library for interacting with SharePoint, supporting CRUD operations, file uploads/downloads, and folder management using WebSocket. The library supports both synchronous and asynchronous operations using promises and async/await.

## Features

- Authenticate with SharePoint using various authentication methods.
- List the contents of a folder (Folders and Files).
- Download files as buffers or to local paths.
- Upload files with support for large file uploads using chunking.
- Create and delete folders.
- Delete files.

## Installation

To install the library, use npm or yarn:

```sh
npm install sp-api-client
```
or
```sh
yarn add sp-api-client
```

## Usage

### Importing the Library

```typescript
import { SharePoint, SharePointOptions } from "sp-api-client";
```

### Initialization

To initialize the SharePoint client, provide the necessary options including base URL, site URL, and authentication information.

```typescript
const options: SharePointOptions = {
  baseUrl: "https://your-sharepoint-site.com",
  siteUrl: "/sites/yoursite",
  loginInfo: {
    username: 'your-username',
    password: 'your-password'
  }
};

async function main() {
  const sharePoint = await SharePoint.create(options);
  // Your further operations
}

main().catch(console.error);
```

### Listing Folder Contents

```typescript
async function listFolderContents() {
  const folderContents = await sharePoint.listFolderContents("/path/to/folder");
  console.log('Folders:', folderContents.folders);
  console.log('Files:', folderContents.files);
}

listFolderContents().catch(console.error);
```

### Downloading Files

#### As a Buffer

```typescript
async function downloadFile() {
  const buffer = await sharePoint.downloadFileAsBuffer("/path/to/file.txt");
  console.log("File Buffer:", buffer);
}

downloadFile().catch(console.error);
```

#### To a Local Path

```typescript
async function downloadFileToLocal() {
  await sharePoint.downloadFileToLocal("/path/to/file.txt", "./local-file.txt");
  console.log("File downloaded to local path.");
}

downloadFileToLocal().catch(console.error);
```

### Uploading Files

#### From a Buffer

```typescript
import * as fs from "fs";

async function uploadFile() {
  const buffer = fs.readFileSync("./local-file.txt");
  await sharePoint.uploadFile("/path/to/folder", "uploaded-file.txt", buffer);
  console.log("File uploaded to SharePoint.");
}

uploadFile().catch(console.error);
```

#### Large File Upload (Chunked)

```typescript
async function uploadLargeFile() {
  const buffer = fs.readFileSync("./large-file.txt");
  await sharePoint.uploadFileFromLargeBuffer("/path/to/folder", "large-uploaded-file.txt", buffer);
  console.log("Large file uploaded to SharePoint.");
}

uploadLargeFile().catch(console.error);
```

### Creating a Folder

```typescript
async function createFolder() {
  await sharePoint.createFolder("/path/to/parent-folder", "new-folder");
  console.log("Folder created.");
}

createFolder().catch(console.error);
```

### Deleting a File

```typescript
async function deleteFile() {
  await sharePoint.deleteFile("/path/to/folder", "file-to-delete.txt");
  console.log("File deleted.");
}

deleteFile().catch(console.error);
```

## API

### SharePoint

#### Methods

- `static create(options: SharePointOptions): Promise<SharePoint>`
  - Initialize the SharePoint client.

- `listFolderContents(folderPath: string): Promise<{ folders: any[], files: any[] }>`
  - Lists the contents of a specified folder.

- `downloadFileAsBuffer(filePath: string): Promise<Buffer>`
  - Downloads a file as a buffer.

- `downloadFileToLocal(filePath: string, localPath: string): Promise<void>`
  - Downloads a file to a local path.

- `uploadFile(folderPath: string, fileName: string, buffer: Buffer): Promise<void>`
  - Uploads a file from a buffer.

- `uploadFileFromLargeBuffer(folderPath: string, fileName: string, buffer: Buffer): Promise<void>`
  - Uploads a large file from a buffer using chunking.

- `createFolder(parentPath: string, folderName: string): Promise<void>`
  - Creates a folder.

- `deleteFile(folderPath: string, fileName: string): Promise<void>`
  - Deletes a file.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request or open an Issue if you have any suggestions or find any bugs.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

For any questions or concerns, please reach out to [meowen0228@gmail.com](mailto:meowen0228@gmail.com).