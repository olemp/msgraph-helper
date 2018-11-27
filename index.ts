import { GraphError } from '@microsoft/microsoft-graph-client';
import { MSGraphClient, MSGraphClientFactory } from '@microsoft/sp-http';

export default class MSGraphHelper {
    private static _graphClient: MSGraphClient;
    public static async Init(msGraphClientFactory: MSGraphClientFactory) {
        this._graphClient = await msGraphClientFactory.getClient();
    }

    // let values: any[] = [];
    // while (true) {
    //     let response: GraphHttpClientResponse = await graphClient.get(url, GraphHttpClient.configurations.v1);
    //     // Check that the request was successful
    //     if (response.ok) {
    //         let result = await response.json();
    //         let nextLink = result["@odata.nextLink"];
    //         // Check if result is single entity or an array of results
    //         if (result.value && result.value.length > 0) {
    //             values.push.apply(values, result.value);
    //         }
    //         result.value = values;
    //         if (nextLink) {
    //             url = result["@odata.nextLink"].replace("https://graph.microsoft.com/", "");
    //         } else {
    //             return result;
    //         }
    //     }
    //     else {
    //         // Reject with the error message
    //         throw new Error(response.statusText);
    //     }
    // }

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
    public static async Get(apiUrl: string, version: string = "v1.0", selectProperties?: Array<string>, filter?: string, top?: number, expand?: string): Promise<any> {
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
            let callback = (error: GraphError, response: any) => {
                if (error) {
                    throw new Error(error.message);
                } else {
                    let nextLink = response["@odata.nextLink"];
                    if (response.value && response.value.length > 0) {
                        values.push(response.value);
                    }
                    if (!nextLink) {
                        return values;
                    }
                }
            };
            await query.get(callback);
        }
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