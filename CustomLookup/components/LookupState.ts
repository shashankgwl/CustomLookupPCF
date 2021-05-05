export interface ILookupState {
    currentItemText?: string,
    currentItemId? : string,
    isModelOpen?: boolean,
    //displayLoading?: boolean,
    lookupData?: ILookupData[],
    lookupColumns?: string[]
}

export interface ILookupData {
    id?: string
}

export interface IDynamicsColumn {
    schemaName: string,
    formattedSchemaName: string,
    displayName: string,
    type: string
}

export interface IJsonContext {
    entityToFetch: string,
    nameAttribute: string,
    lookupAttributeOnPage: string,
    entityColumns: IDynamicsColumn[]
}

export interface IPageContext {
    webResourceURL: string
}