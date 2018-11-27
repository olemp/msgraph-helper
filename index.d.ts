export default class MSGraphHelper {
    private static _graphClient;
    static Init(msGraphClientFactory: any): Promise<void>;
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
    static Get(apiUrl: string, version?: string, selectProperties?: Array<string>, filter?: string, top?: number, expand?: Array<string>): Promise<any>;
    /**
     * Patch
     *
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     * @param {any} content Content
     */
    static Patch(apiUrl: string, version: string, content: any): Promise<any>;
    /**
     * Post
     *
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     * @param {any} content Content
     */
    static Post(apiUrl: string, version: string, content: any): Promise<any>;
    /**
     * Delete
     *
     * @param {string} apiUrl API url
     * @param {string} version Version (default to v1.0)
     */
    static Delete(apiUrl: string, version?: string): Promise<any>;
}
