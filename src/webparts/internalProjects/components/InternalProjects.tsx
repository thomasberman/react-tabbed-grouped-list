import * as React from 'react';
import styles from './InternalProjects.module.scss';
import { IInternalProjectsProps } from './IInternalProjectsProps';
import { groupBy, findIndex } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import { Pivot, PivotItem, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { GroupedList, IGroup, IGroupDividerProps } from 'office-ui-fabric-react/lib/components/GroupedList/index';
import { GroupHeader } from 'office-ui-fabric-react/lib/components/GroupedList/GroupHeader';
import { SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';

export interface IInternalProjectsState {
  listItemsGroupedByOffice: _.Dictionary<any[]>;
}

export default class InternalProjects extends React.Component<IInternalProjectsProps, IInternalProjectsState> {

  public componentDidMount(): void {
    this._getAndGroupInternalProjectItems().then(res => {
      this.setState({
        listItemsGroupedByOffice: res
      });
    });
  }

  public render(): React.ReactElement<IInternalProjectsProps> {
    return (
      <div className={styles.internalProjects}>
        <Pivot linkSize={PivotLinkSize.large}>
          {this.state && this._generateJSXMarkup(this.state.listItemsGroupedByOffice)}
        </Pivot>
      </div>
    );
  }

  private _getAndGroupInternalProjectItems(): Promise<_.Dictionary<any[]>> {
    return sp.web.lists.getByTitle("Internal Projects").items.get().then(res => groupBy(res, 'Office_x0020_Location'));
  }

  private _generateJSXMarkup(listItemsGroupedByOffice: _.Dictionary<any[]>): JSX.Element[] {
    let JSXMarkup = [];

    Object.keys(listItemsGroupedByOffice).forEach((location, index) => {
      const sortedLocationArr: any[] = listItemsGroupedByOffice[location].sort((a, b) => (a.Department < b.Department) ? -1 : (a.Department > b.Department) ? 1 : 0);
      const departmentGroups: IGroup[] = this._generateIGroupsFromArray(sortedLocationArr);

      let children = [];

      children.push(
        <GroupedList
          items={listItemsGroupedByOffice[location]}
          groupProps={{
            onRenderHeader: this._onRenderHeader
          }}
          groups={departmentGroups}
          onRenderCell={this._onRenderCell}
          selectionMode={SelectionMode.none}
        />
      );
      JSXMarkup.push(<PivotItem linkText={location} key={index}>{children}</PivotItem>);
    });
    return JSXMarkup;
  }

  private _generateIGroupsFromArray(sortedOfficeLocationItems: any[]): IGroup[] {
    const groupedByDepartments = groupBy(sortedOfficeLocationItems, (i: any) => i.Department);
    let groups: IGroup[] = [];

    for (const x in groupedByDepartments) {
      groups.push({
        name: x,
        key: x,
        startIndex: findIndex(sortedOfficeLocationItems, (i: any) => i.Department === x),
        count: groupedByDepartments[x].length,
        isCollapsed: true
      });
    }
    return groups;
  }

  private _onRenderHeader(props: IGroupDividerProps): JSX.Element {
    const onToggleSelectGroup = () => {
      props.onToggleCollapse(props.group);
    };
    return (
      <GroupHeader {...props} onToggleSelectGroup={onToggleSelectGroup} />
    );
  }

  private _onRenderCell(nestingDepth: number, item: any, itemIndex: number): JSX.Element {
    return <div className={styles.subTitle}>{item.Title}</div>;
  }

}
