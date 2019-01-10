import { GraphError } from '@microsoft/microsoft-graph-client';
import { MSGraphClient } from '@microsoft/sp-http';
import { Logger, LogLevel, ConsoleListener } from '@pnp/logging';

export default class MSGraphHelper {
    private static _graphClient: MSGraphClient;
    public static async Init(msGraphClientFactory) {
        this._graphClient = await msGraphClientFactory.getClient();
        Logger.subscribe(new ConsoleListener());
        Logger.activeLogLevel = LogLevel.Info;
    }

    /**
     * Get
     * 
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     * @param {Array<string>} selectProperties Select properties
     * @param {string} filter Filter
     * @param {number} top Number of items to retrieve
     * @param {Array<string>} expand Expand
     */
    public static async Get(apiUrl: string, version: string = "v1.0", selectProperties?: Array<string>, filter?: string, top?: number, expand?: Array<string>): Promise<any> {
        try {
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
            Logger.log({ message: `(MSGraphHelper) Get`, data: { urlComponents: query.urlComponents }, level: LogLevel.Info });

            let response = await query.get();
            if (response.value && response.value.length > 0) {
                values.push(...response.value);
            }
            let nextLink = response["@odata.nextLink"];
            if (nextLink) {
                while (true) {
                    try {
                        query.parsePath(nextLink);
                        response = await query.get();
                        nextLink = response["@odata.nextLink"];
                        if (response.value && response.value.length > 0) {
                            values.push(...response.value);
                        }
                        if (!nextLink) {
                            break;
                        }
                    } catch (error) {
                        throw error;
                    }
                }
            }
            return values;
        } catch (error) {
            throw error;
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