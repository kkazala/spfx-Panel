declare interface IPanelCommandSetStrings {
    Command1: string;
    Command2: string;

    titleTravelGuidelines:string;
    titleTravelReport:string;

    lblRefreshing:string;
    lblItemUpdate_OK:string
    lblItemUpdate_Err:string

    lblConfirm:string
    lblYes:string;
    lblNo:string;
    lblPageWillRefresh:string

    btnSubmit:string;

    htmlInfo:string;
}

declare module 'PanelCommandSetStrings' {
  const strings: IPanelCommandSetStrings;
  export = strings;
}
