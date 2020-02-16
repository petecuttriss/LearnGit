declare interface IMyTeamsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
}

declare module 'MyTeamsWebPartStrings' {
  const strings: IMyTeamsWebPartStrings;
  export = strings;
}
