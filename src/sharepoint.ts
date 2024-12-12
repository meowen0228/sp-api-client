import * as fs from "fs";
import axios, { AxiosRequestConfig } from "axios";
import * as spauth from "node-sp-auth";
import * as moment from "moment";
import { v4 } from "uuid";

import type { IAuthOptions } from "node-sp-auth";

interface SharePointOptions {
  baseUrl: string;
  siteUrl: string;
  loginInfo: IAuthOptions;
}

export class SharePoint {
  private baseSiteUrl: string;
  private apiUrl: string;
  private requestOpts: AxiosRequestConfig;

  constructor(private options: SharePointOptions) {
    const { baseUrl, siteUrl } = options;
    this.baseSiteUrl = `${baseUrl}${siteUrl}`;
    this.apiUrl = `${this.baseSiteUrl}/_api/web`;
    this.requestOpts = {
      headers: {},
    };
  }

  static async create(options: SharePointOptions): Promise<SharePoint> {
    const sharePoint = new SharePoint(options);
    await sharePoint.getAuth();
    return sharePoint;
  }

  private async getAuth(): Promise<void> {
    const { loginInfo } = this.options;
    const authResponse = await spauth.getAuth(this.baseSiteUrl, loginInfo);
    const headers = { ...authResponse.headers };
    headers["Authorization"] = `Bearer ${authResponse.headers.Cookie}`;
    headers["Accept"] = "application/json;odata=verbose";
    headers["Content-Type"] = "application/json";

    const contextInfoResponse = await axios.post(
      `${this.baseSiteUrl}/_api/contextinfo`,
      {},
      { headers }
    );

    headers["X-RequestDigest"] = contextInfoResponse.data.d.GetContextWebInformation.FormDigestValue;
    this.requestOpts.headers = headers;
  }

  public async listFolderContents(folderPath: string) {
    const path = this.formatPath(folderPath);
    const url = `${this.apiUrl}/GetFolderByServerRelativeUrl('${path}')?$expand=Folders,Files`;
    const response = await axios.get(url, this.requestOpts);
    const sortFunc = (a: any, b: any) => moment(b.TimeCreated).diff(moment(a.TimeCreated));

    return {
      folders: response.data.d.Folders.results.sort(sortFunc),
      files: response.data.d.Files.results.sort(sortFunc),
    };
  }

  public async downloadFileAsBuffer(filePath: string): Promise<Buffer> {
    const option = { ...this.requestOpts };
    option.responseType = "arraybuffer";
    const path = this.formatPath(filePath);
    const url = `${this.apiUrl}/GetFileByServerRelativeUrl('${path}')/$value`;
    const response = await axios.get(url, option);
    return Buffer.from(response.data);
  }

  public async downloadFileToLocal(filePath: string, localPath: string): Promise<void> {
    const buffer = await this.downloadFileAsBuffer(filePath);
    fs.writeFileSync(localPath, Buffer.from(buffer));
  }

  public async uploadFile(folderPath: string, fileName: string, buffer: Buffer): Promise<void> {
    this.validateName(fileName);
    const path = this.formatPath(folderPath);
    const url = `${this.apiUrl}/GetFolderByServerRelativePath(DecodedUrl='${path}')/Files/AddUsingPath(DecodedUrl='${fileName}',overwrite=true)`;

    await axios.post(url, buffer, {
      ...this.requestOpts,
      maxContentLength: Infinity,
      maxBodyLength: Infinity,
    });
  }

  public async uploadFileFromLargeBuffer(folderPath: string, fileName: string, buffer: Buffer): Promise<void> {
    const uploadId = v4();
    const path = this.formatPath(folderPath);
    const option = this.requestOpts;
    option.method = "post";

    console.log(`share point: start upload ${fileName}`);
    option.url = `${this.apiUrl}/GetFolderByServerRelativePath(DecodedUrl='${path}')/Files/AddStubUsingPath(DecodedUrl='${fileName}')/StartUploadFile(uploadId='${uploadId}')`;
    await axios(option);

    const chunkSize = 100 * 1024 * 1024;
    let finalOffset = 0;
    for (let offset = 0; offset < buffer.length; offset += chunkSize) {
      console.log(
        `share point: uploading ${fileName} buffer: ${offset} to ${offset +
        chunkSize}`
      );
      option.url = `${this.apiUrl}/GetFileByServerRelativePath(DecodedUrl='${path}/${fileName}')/ContinueUpload(uploadId='${uploadId}',fileOffset='${offset}')`;
      const chunk = buffer.slice(offset, offset + chunkSize);
      finalOffset = chunk.length;
      option.data = chunk;
      option.maxContentLength = Infinity;
      option.maxBodyLength = Infinity;
      await axios(option);
    }

    option.url = `${this.apiUrl
      }/GetFileByServerRelativePath(DecodedUrl='${path}/${fileName}')/FinishUpload(uploadId='${uploadId}',fileOffset='${buffer.length -
      finalOffset}')`;
    await axios(option);
    console.log(`share point: upload done ${fileName}`);
  }

  public async createFolder(parentPath: string, folderName: string): Promise<void> {
    this.validateName(folderName);
    const path = this.formatPath(parentPath);
    const url = `${this.apiUrl}/folders/AddUsingPath(DecodedUrl='${path}/${folderName}',overwrite=false)`;

    await axios.post(url, {}, this.requestOpts);
  }

  public async deleteFile(folderPath: string, fileName: string): Promise<void> {
    this.validateName(fileName);
    const path = this.formatPath(folderPath);
    const url = `${this.apiUrl}/GetFileByServerRelativeUrl('${path}/${fileName}')`;

    await axios.delete(url, this.requestOpts);
  }

  private validateName(name: string): void {
    const invalidChars = /["*:<|>?\\/]/;
    if (invalidChars.test(name)) {
      throw new Error(`Invalid name: ${name} contains forbidden characters.`);
    }
  }

  private formatPath(path: string): string {
    return encodeURI(path.startsWith(this.options.siteUrl) ? path : `${this.options.siteUrl}${path}`);
  }
}
