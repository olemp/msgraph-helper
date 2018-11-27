import { GraphError } from '@microsoft/microsoft-graph-client';
import { MSGraphClient } from '@microsoft/sp-http';

export default class MSGraphHelper {
    private static _graphClient: MSGraphClient;
    public static async Init(msGraphClientFactory) {
        this._graphClient = await msGraphClientFactory.getClient();
    }

    /**
     * Get
     * 
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     * @param {Array<string>} selectProperties Select properties
     * @param {string} filter Filter
     * @param {number} top Number of items to retrieve
     * @param {string} expand Expand
     */
    public static Get(apiUrl: string, version: string = "v1.0", selectProperties?: Array<string>, filter?: string, top?: number, expand?: string): Promise<any> {
        return new Promise<any>(async (resolve, reject) => {
            let values = [];
            let query = this._graphClient.api(apiUrl).version(version);
            if (selectProperties && selectProperties.length > 0) {
                query = query.select(selectProperties);
            }
            if (filter && filter.length > 0) {
                query = query.filter(filter);
            }
            if (top) {
                query = query.top(top);
            }
            if (expand) {
                query = query.expand(expand);
            }

            while (true) {
                await query.get((error: GraphError, response: any) => {
                    if (error) {
                        reject(error);
                    } else {
                        let nextLink = response["@odata.nextLink"];
                        if (response.value && response.value.length > 0) {
                            values.push(response.value);
                        }
                        if (!nextLink) {
                            resolve(values);
                        }
                    }
                });
            }
        });
    }

    /**
     * Patch
     * 
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     * @param {any} content Content
     */
    public static async Patch(apiUrl: string, version: string = "v1.0", content: any): Promise<any> {
        var p = new Promise<string>(async (resolve, reject) => {
            if (typeof (content) === "object") {
                content = JSON.stringify(content);
            }

            let query = this._graphClient.api(apiUrl).version(version);
            let callback = (error: GraphError, _response: any, rawResponse?: any) => {
                if (error) {
                    reject(error);
                } else {
                    resolve();
                }
            };
            await query.update(content, callback);
        });
        return p;
    }

    /**
     * Post
     * 
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     * @param {any} content Content
     */
    public static async Post(apiUrl: string, version: string = "v1.0", content: any): Promise<any> {
        var p = new Promise<string>(async (resolve, reject) => {
            if (typeof (content) === "object") {
                content = JSON.stringify(content);
            }

            let query = this._graphClient.api(apiUrl).version(version);
            let callback = (error: GraphError, response: any, rawResponse?: any) => {
                if (error) {
                    reject(error);
                } else {
                    resolve(response);
                }
            };
            await query.post(content, callback);
        });
        return p;
    }

    /**
     * Delete
     * 
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     */
    public static async Delete(apiUrl: string, version: string = "v1.0"): Promise<any> {
        var p = new Promise<string>(async (resolve, reject) => {
            let query = this._graphClient.api(apiUrl).version(version);
            let callback = (error: GraphError, response: any, rawResponse?: any) => {
                if (error) {
                    reject(error);
                } else {
                    resolve(response);
                }
            };
            await query.delete(callback);
        });
        return p;
    }
}