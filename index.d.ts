import { WebPartContext } from '@microsoft/sp-webpart-base';
export default class MSGraphHelper {
    private static _graphClient;
    static Init(context: WebPartContext): Promise<void>;
    static Get(apiUrl: string, version?: string, selectProperties?: string[], filter?: string): Promise<any>;
    static Patch(apiUrl: string, version: string, content: any): Promise<any>;
    static Post(apiUrl: string, version: string, content: any): Promise<any>;
    static Delete(apiUrl: string, version?: string): Promise<any>;
}
