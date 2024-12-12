# sp-api-client 功能指南

sp-api-client 讓您可以輕鬆操作檔案及資料夾。以下提供功能概述與操作範例，方便使用者理解如何使用此模組。

## 功能清單

### 1. 列出資料夾內容

顯示資料夾中的所有檔案與子資料夾。

#### 範例

```typescript
import { SharePoint } from "./dist/sharepoint";

const sharePoint = await SharePoint.create({
  baseUrl: "baseUrl";
  siteUrl: "siteUrl";
  loginInfo: {
    userusername: "userusername";
    password: "password";
  };
})
const contents = await sharePoint.listFolderContents("/test-folder");
console.log(contents);
// result：
// {
//   folders: ["folder1", "folder2"],
//   files: ["file1.txt", "file2.txt"]
// }
```

### 2. 下載檔案

從 SharePoint 下載檔案到本地端。

```typescript
await sharePoint.downloadFileToLocal(
  "/test-folder/test.txt",
  "./local-test.txt"
);
console.log("檔案下載完成！");
```

### 3. 上傳檔案

將本地端檔案上傳到 SharePoint 指定資料夾。

```typescript
const buffer = Buffer.from("檔案內容");
await sharePoint.uploadFile("/test-folder", "new-file.txt", buffer);
console.log("檔案上傳成功！");
```

### 4. 建立資料夾

在 SharePoint 上建立新資料夾。

```typescript
await sharePoint.createFolder("/parent-folder", "new-folder");
console.log("資料夾建立完成！");
```

### 5. 刪除檔案

刪除 SharePoint 中的指定檔案。

```typescript
await sharePoint.deleteFile("/test-folder", "old-file.txt");
console.log("檔案已刪除！");
```

## 注意事項

- 請確認您使用的 SharePoint 帳戶擁有操作權限。
- 測試時確保 `options` 中的站點 URL 和登入資訊正確無誤。

## 結語

本模組提供基本的檔案與資料夾操作功能，適用於開發和測試環境。如果您有任何建議或發現問題，歡迎提交回饋或貢獻代碼！
