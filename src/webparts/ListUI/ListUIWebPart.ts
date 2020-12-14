import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneDropdown, IPropertyPaneDropdownOption, PropertyPaneCheckbox } from '@microsoft/sp-webpart-base';
import * as strings from 'ListUIWebPartStrings';
import ListUI from './components/ListUI';
import { IListUIProps } from './components/IListUIProps';
import { IListColumn } from './services/IListColumn';
import { ListService } from './services/ListService';
import Constants from './components/Constants';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldOrder } from '@pnp/spfx-property-controls/lib/PropertyFieldOrder';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';

export interface IListUIWebPartProps {
  formPageName: string;
  listName: string;
  pageSize: number;
  listDisplayType: string;
  listColumnsFetched: boolean;
  listColumnsAndTypes: IItemProp[];
  listColumnsSelected: string[];
}

export interface IItemProp { key: string; text: string; displayName: string; }

export default class ListUIWebPart extends BaseClientSideWebPart<IListUIWebPartProps> {

  protected onInit(): Promise<void> {

    // hardcoded list name by purpose
    this.properties.listName = Constants._ListName_customList;
    // hardcoded list name by purpose

    // default value is 1 - all entries
    this.properties.listDisplayType = this.properties.listDisplayType === null || this.properties.listDisplayType === undefined ? "1" : this.properties.listDisplayType;

    // init to empty list to avoid first loading issue
    //this.properties.listColumnsAndTypes = this.properties.listColumnsAndTypes === null || this.properties.listColumnsAndTypes === undefined ? [] : this.properties.listColumnsAndTypes;
    //this.properties.listColumnsSelected = this.properties.listColumnsSelected === null || this.properties.listColumnsSelected === undefined ? [] : this.properties.listColumnsSelected;
    //this.properties.listColumnsFetched = false;

    this.configureWebPart = this.configureWebPart.bind(this);
    this.selectedColumns = this.selectedColumns.bind(this);
    return super.onInit();
  }

  protected onPropertyPaneConfigurationStart(): void { this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'configuration changes. Please refresh the page when done.'); }
  protected onPropertyPaneConfigurationEnd(): void { this.context.statusRenderer.clearLoadingIndicator(this.domElement); }

  public render(): void {
    const element: React.ReactElement<IListUIProps> = React.createElement(
      ListUI,
      {
        context: this.context,
        currentUserLoginName: this.context.pageContext.user.loginName,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        formPageName: this.properties.formPageName,
        listName: this.properties.listName,
        needsConfiguration: this.needsConfiguration(),
        configureWebPart: this.configureWebPart,
        displayMode: this.displayMode,
        selectedColumns: this.selectedColumns(),
        pageSize: this.properties.pageSize,
        listDisplayType: this.properties.listDisplayType
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected get disableReactivePropertyChanges(): boolean { return true; }
  private configureWebPart(): void { this.context.propertyPane.open(); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (!this.properties.listColumnsFetched) {
      this.loadColumns().then((response) => {
        this.properties.listColumnsFetched = true;
        this.context.propertyPane.refresh();
        // this.onDispose(); // Dispose is crucial - async method needs to finish before showing property pane!
      });
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('formPageName', {
                  label: strings.FormPageNameFieldLabel
                }),
                PropertyPaneDropdown('listDisplayType', {
                  label: strings.ListDisplayTypeFieldLabel,
                  options: [
                    { key: '1', text: 'All items' },
                    { key: '2', text: 'Current user is author' },
                    { key: '3', text: 'State - Approved' },
                  ]
                }),
                PropertyPaneDropdown('pageSize', {
                  label: strings.PageSizeFieldLabel,
                  options: [
                    { key: '2', text: '2' },
                    { key: '10', text: '10' },
                    { key: '25', text: '25' },
                    { key: '50', text: '50' },
                    { key: '100', text: '100' }
                  ]
                }),
                PropertyFieldMultiSelect('listColumnsSelected', {
                  key: 'listColumnsSelected',
                  label: strings.ColumnFieldLabel,
                  options: this.properties.listColumnsAndTypes && this.properties.listColumnsAndTypes.map((item: IItemProp) => { return { key: item.key, text: item.displayName + " (" + item.key + ")" }; }),
                  selectedKeys: this.properties.listColumnsSelected
                }),
                PropertyFieldOrder("orderedItems", {
                  label: strings.OrderedItemsFieldLabel,
                  key: "orderedItems",
                  items: this.properties.listColumnsSelected,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private needsConfiguration(): boolean {
    return this.properties.listName === null ||
      this.properties.listName === undefined ||
      this.properties.listName.trim().length === 0 ||
      this.properties.listColumnsSelected === null ||
      this.properties.listColumnsSelected === undefined ||
      this.properties.listColumnsSelected.length === 0;
  }

  private loadColumns(): Promise<IItemProp[]> {
    const dataService = new ListService(this.context);
    return new Promise<IItemProp[]>(resolve => {
      dataService.getColumns(this.properties.listName)
        .then((response) => {
          var options: IItemProp[] = [];
          this.properties.listColumnsAndTypes = [];
          response.forEach((column: IListColumn) => {
            options.push({ "key": column.StaticName, "text": column.Title, "displayName": column.Title });
            this.properties.listColumnsAndTypes.push({ "key": column.StaticName, "text": column.TypeDisplayName, "displayName": column.Title });
            // StaticName - internal column name
            // Title - display name of column
            // TypeDisplayName - data type of column 'Person or Group', etc.
          });
          resolve(options);
        });
    });
  }

  private selectedColumns(): IItemProp[] {
    if (this.properties.listColumnsAndTypes === null || this.properties.listColumnsAndTypes === undefined || this.properties.listColumnsAndTypes.length === 0) {
      return [];
    }
    else {
      var resultColumns: IItemProp[] = [];
      var listColumnsAndTypes: IItemProp[] = this.properties.listColumnsAndTypes;
      this.properties.listColumnsSelected.forEach(function (value) {
        var item = listColumnsAndTypes.filter(obj => obj.key === value)[0];
        resultColumns.push(item);
      });
      return resultColumns;
    }
  }
}