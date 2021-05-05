import * as React from 'react'
import { ShimmeredDetailsList, Fabric, StackItem, Link, PrimaryButton, Modal, IColumn, Text, SelectionMode, TextField } from 'office-ui-fabric-react/lib'
import { FontWeights, IconButton, Stack } from '@fluentui/react'
import { mergeStyleSets, getTheme } from 'office-ui-fabric-react';
import { IJsonContext, IDynamicsColumn, ILookupData, ILookupState, IPageContext } from './LookupState';
import { initializeIcons } from '@uifabric/icons';

export default function CustomLookupCore(props: IPageContext) {

    const [lookupState, setLookupState] = React.useState<ILookupState>({});
    const [cachedLookupData, setCachedLookupData] = React.useState<any[]>([]);
    const [isBusy, setIsBusy] = React.useState(true);
    const [jsonContext, setJsonContext] = React.useState<IJsonContext>({
        entityColumns: [],
        nameAttribute: "name",
        entityToFetch: '',
        lookupAttributeOnPage: ''
    });


    initializeIcons();

    const theme = getTheme();
    const contentStyles = mergeStyleSets({
        container: {
            display: 'flex',
            flexFlow: 'column nowrap',
            alignItems: 'stretch',
            width: '80%',
            height: '60%'
        },
        header: [

            theme.fonts.xLargePlus,
            {
                flex: '1 1 auto',
                borderTop: `4px solid ${theme.palette.themePrimary}`,
                color: theme.palette.neutralPrimary,
                display: 'flex',
                alignItems: 'center',
                fontWeight: FontWeights.semibold,
                padding: '12px 12px 14px 24px',
            },
        ],
        body: {
            flex: '4 4 auto',
            padding: '0 24px 24px 24px',
            overflowY: 'hidden',
            selectors: {
                p: { margin: '14px 0' },
                'p:first-child': { marginTop: 0 },
                'p:last-child': { marginBottom: 0 },
            },
        },
    });



    async function getWebResource(url: string) {
        const response = await fetch(url)
        return response.text();
    }

    const getFormattedColumn = (schema: { displayName: string; schemaName: string; type: string; }) => {
        switch (schema.type) {
            case 'lookup':
                return '_' + schema.schemaName + '_value'
                break;
            default:
                return schema.schemaName;
        }
        return ''
    }

    const getDynamicsColumn = (jsonCol: any): IDynamicsColumn[] => {
        const obj: IDynamicsColumn[] = []
        jsonCol.map((col: { displayName: any; schemaName: any; type: any; }) => obj.push({
            displayName: col.displayName,
            schemaName: col.schemaName,
            type: col.type,
            formattedSchemaName: getFormattedColumn(col)
        }))

        return obj
    }

    React.useEffect(() => {

        getWebResource(props.webResourceURL).then(json => {
            const jobject = JSON.parse(json);
            const dataContext: IJsonContext = {
                entityToFetch: jobject.entityToFetch,
                nameAttribute: jobject.nameAttribute,
                entityColumns: getDynamicsColumn(jobject.columns),
                lookupAttributeOnPage: jobject.lookupAttributeOnPage
            };


            setJsonContext(dataContext);
            const pageAttribute = Xrm.Page.data.entity.attributes.get(dataContext.lookupAttributeOnPage).getValue();
            
            const state: ILookupState =
            {
                currentItemText: pageAttribute ? pageAttribute[0].name : '',
                currentItemId: pageAttribute ? pageAttribute[0].id : '',
                isModelOpen: false,
                lookupColumns: Array.from(dataContext.entityColumns, col => col.displayName)
                //..lookupData: []
            }


            //Xrm.Utility.alertDialog(state.lookupColumns!.toString(), () => { })

            setLookupState(state);


        }).catch(err => {
            localAlert(err);
        });

    }, [props])

    function localAlert(msg?: string) {
        Xrm.Utility.alertDialog(msg!, () => { });
    }

    function timeout(ms: number) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    async function buildQuery() {
        //Xrm.WebApi.online.retrieveMultipleRecords("contact", "?$select=coke_kouserid,coke_name, coke_regionid_value")
        let query = "?$select="
        jsonContext.entityColumns.map(item => {
            query += item.formattedSchemaName + ','
        })

        if (query.endsWith(',')) {
            query = query.substr(0, query.lastIndexOf(',')) //+ '&$top=2'
        }

        //localAlert(query);

        return query;
    }

    //id: records.entities[i][jsonContext.entityToFetch + 'id'],
    //text: records.entities[i][jsonContext.entityColumns[j].formattedSchemaName + "@OData.Community.Display.V1.FormattedValue"]

    const getDataFromCE = async (query: string) => {
        const data: ILookupData[] = []
        let tempRow: ILookupData = {}
        const records = await Xrm.WebApi.online.retrieveMultipleRecords(jsonContext.entityToFetch, query)
        //localAlert(records.entities.length.toString());
        for (let i = 0; i < records.entities.length; i++) {
            tempRow = {}
            tempRow.id = records.entities[i][jsonContext.entityToFetch + 'id']
            for (let j = 0; j < jsonContext.entityColumns.length; j++) {

                if (jsonContext.entityColumns[j].type === 'lookup') {
                    Object.defineProperty(tempRow, jsonContext.entityColumns[j].schemaName,
                        {
                            value: records.entities[i][jsonContext.entityColumns[j].formattedSchemaName + "@OData.Community.Display.V1.FormattedValue"],
                            writable: true,
                            configurable: true,
                            enumerable: true
                        });
                }

                else {
                    Object.defineProperty(tempRow, jsonContext.entityColumns[j].schemaName,
                        {
                            value: Object.values(tempRow)[j] = records.entities[i][jsonContext.entityColumns[j].formattedSchemaName],
                            writable: true,
                            configurable: true,
                            enumerable: true
                        });


                }
            }

            data.push(tempRow);

        }
        console.log(data);
        return data;
    }

    const onDialogOpen = async () => {
        setIsBusy(true);
        setLookupState({
            currentItemText: lookupState.currentItemText,
            isModelOpen: true,
            lookupData: [],
            lookupColumns: lookupState.lookupColumns
        })

        const lookupData = await getDataFromCE(await buildQuery());

        const state: ILookupState =
        {
            isModelOpen: true,
            lookupData: lookupData,
            lookupColumns: lookupState.lookupColumns,
            currentItemText: lookupState.currentItemText
        }

        setLookupState(state);
        setCachedLookupData(lookupData || [])
        setIsBusy(false);
    }

    function getListColumns() {
        const cols: IColumn[] = [];
        cols.push({
            key: "id",
            minWidth: 80,
            maxWidth: 180,
            name: "Select",
            isResizable: true,
            isCollapsible: true,
            data: 'string'

        });
        if (!lookupState.lookupColumns)
            return [];
        for (let i = 0; i < jsonContext.entityColumns!.length; i++) {
            cols.push({
                key: jsonContext.entityColumns[i].schemaName,
                minWidth: 80,
                maxWidth: 180,
                name: lookupState.lookupColumns![i],
                isResizable: true,
                isCollapsible: true,
                data: 'string'

            })
        }

        return cols;
    }

    //const getData = (): ILookupData[] => {
    //    var data: ILookupData[] = []


    //    for (var i = 0; i < 50; i++) {
    //        data.push({
    //            id: i.toString(),
    //            text: "text" + i
    //        });
    //    }

    //    return data;
    //}


    const onTextFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, tx?: string): void => {
        var element: any = ev
        try {
            const filtered = tx ? cachedLookupData.filter(
                item => item[element.target.name] &&
                    item[element.target.name].toLowerCase().indexOf(tx.toLowerCase()) > -1) : cachedLookupData

            //localAlert(filtered.length.toString())

            setLookupState({
                currentItemText: lookupState.currentItemText,
                lookupData: filtered,
                isModelOpen: true,
                lookupColumns: lookupState.lookupColumns
            })

        } catch (error) {
            localAlert(error)
        }

    }

    function _renderItemColumn(item?: ILookupData, index?: number, column?: IColumn) {
        if (column?.key === "id") {
            return (

                <Link onClick={() => {
                    var lookupValue = new Array();
                    lookupValue[0] = new Object();
                    lookupValue[0].id = item?.id; // GUID of the lookup id
                    lookupValue[0].name = item![jsonContext.nameAttribute as keyof ILookupData]; // Name of the lookup
                    lookupValue[0].entityType = jsonContext.entityToFetch; //Entity Type of the lookup entity
                    Xrm.Page.getAttribute(jsonContext.lookupAttributeOnPage).setValue(lookupValue); // You need to replace the lookup field 
                    setLookupState(
                        {
                            currentItemText: item![jsonContext.nameAttribute as keyof ILookupData],
                            isModelOpen: false,
                            lookupColumns: lookupState.lookupColumns,
                            currentItemId: item?.id

                        })
                }}>Select</Link>
            )
        }
        return (
            <Text>{item![column!.key as keyof ILookupData]}</Text>
        )
    }

    const onDialogDismiss = () => {
        const state: ILookupState =
        {
            isModelOpen: false,
            currentItemText: lookupState.currentItemText,
            lookupColumns: lookupState.lookupColumns
        }

        setLookupState(state);
    }

    return (
        <Fabric>
            <Stack horizontal style={{ border: 2 }}>
                <StackItem>
                    <Link onClick={() => {
                        Xrm.Utility.openEntityForm(jsonContext.entityToFetch, lookupState.currentItemId, undefined, { openInNewWindow: true });
                    }}>{lookupState.currentItemText}  </Link >
                </StackItem>

                <StackItem style={{ marginLeft: 30 }}>
                    <IconButton iconProps={{ iconName: "Search" }} onClick={onDialogOpen} />
                </StackItem>
            </Stack>

            <Modal
                isOpen={lookupState.isModelOpen}
                containerClassName={contentStyles.container}
                onDismiss={onDialogDismiss}>

                <Stack horizontal tokens={{ childrenGap: 5 }}>
                    {
                        jsonContext.entityColumns?.map((columnName) => {
                            return (
                                <Stack.Item>
                                    <TextField name={columnName.schemaName} disabled={isBusy} label={columnName.displayName + ' Filter'} onChange={onTextFilter} />
                                </Stack.Item>
                            )
                        })
                    }

                    <Stack.Item align="end" tokens={{ margin: "0,0,0,50" }}>
                        <PrimaryButton onClick={() => {
                            Xrm.Utility.openEntityForm(jsonContext.entityToFetch, undefined, undefined, { openInNewWindow: true });
                        }} marginWidth={25} text={"Create new " + jsonContext.entityToFetch} />
                    </Stack.Item>
                </Stack>


                <div>
                    <ShimmeredDetailsList
                        selectionMode={SelectionMode.none}
                        compact
                        columns={getListColumns()}
                        items={lookupState.lookupData ? lookupState.lookupData : []}
                        onRenderItemColumn={_renderItemColumn}
                        enableShimmer={isBusy}
                    /></div>
            </Modal>
        </Fabric>
    )
}