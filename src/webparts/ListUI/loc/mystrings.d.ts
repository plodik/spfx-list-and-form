declare interface IListUIWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListNameFieldLabel: string;
  ListDisplayTypeFieldLabel: string;
  ColumnFieldLabel: string;
  PageSizeFieldLabel: string;
  FormPageNameFieldLabel: string;
  OrderedItemsFieldLabel: string;
}

declare module 'ListUIWebPartStrings' {
  const strings: IListUIWebPartStrings;
  export = strings;
}
