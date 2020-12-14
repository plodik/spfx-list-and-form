declare interface IFormUIWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ListPageNameFieldLabel: string;
}

declare module 'FormUIWebPartStrings' {
  const strings: IFormUIWebPartStrings;
  export = strings;
}
