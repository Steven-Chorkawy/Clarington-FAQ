import * as React from 'react';
import { IFaqAccordionProps } from './IFaqAccordionProps';
import { getSiteSP } from '../../../pnpjs-config';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import "@pnp/sp/items/get-all";
import "@pnp/sp/taxonomy";
import { ExpansionPanel, ExpansionPanelActionEvent, ExpansionPanelContent } from '@progress/kendo-react-layout';
import { Reveal } from '@progress/kendo-react-animation';
import "@pnp/sp/taxonomy";
import { Link, MessageBar, MessageBarType, SearchBox } from 'office-ui-fabric-react';
import { filterBy } from '@progress/kendo-data-query';
import { PermissionKind } from "@pnp/sp/security";

const MOC_ORG_GROUP_ID = '4026b60c-6222-432f-b07d-89c2396e8e64';
const DEPARTMENT_TERM_SET_ID = '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f';

export default class FaqAccordion extends React.Component<IFaqAccordionProps, any> {

  constructor(props: any) {
    super(props);
    this.state = {
      items: undefined,
    };
    this._queryList().then().catch(reason => console.error(reason));

    // Check users permissions. 
    getSiteSP().web.lists.getByTitle('FAQ').currentUserHasPermissions(PermissionKind.EditListItems).then((value) => {
      this.setState({
        canUserEditListItems: value
      });
    }).catch(reason => console.error(reason));
  }

  // TODO: This method should query any managed metadata field not just a department field.
  private async _queryDepartmentName(termID: string): Promise<void> {
    if (!termID)
      return;

    // check if state has already been set. 
    if (!this.state[termID]) {
      const res = await getSiteSP().termStore.groups.getById(MOC_ORG_GROUP_ID).sets.getById(DEPARTMENT_TERM_SET_ID).getTermById(termID)();

      if (res.labels) {
        if (res.labels.length > 0) {
          // There should only ever be one item in this array so just get the first one.
          this.setState({ [termID]: res.labels[0].name });

          // Replace the Department object with a simple string of the department name.  
          // This will make searching and rendering this info much easier.
          this.state.items.map((item: any, index: number) => {
            if (item[this.props.subtitleFieldName]?.TermGuid === termID) {
              // ? What is DepartmentDate used for?
              // item.DepartmentDate = { ...item.Department }
              item[this.props.subtitleFieldName] = res.labels[0].name
            }
          });
        }
      }
    }
  }

  private async _queryList(): Promise<void> {
    const listItems = await getSiteSP().web.lists.getByTitle(this.props.listName).items.getAll();

    listItems.forEach(item => {
      this._queryDepartmentName(item[this.props.subtitleFieldName]?.TermGuid);
    });

    // Sort by Created date.  Newest to Oldest.
    const sortedList = listItems.sort((p1, p2) => (p1.Created < p2.Created) ? 1 : (p1.Created > p2.Created) ? -1 : 0);

    console.log('Items');
    console.log(sortedList);
    this.setState({
      items: sortedList,    // items that will be rendered. 
      allItems: sortedList  // All items regardless of current filters.
    });
  }

  componentDidUpdate(prevProps: Readonly<IFaqAccordionProps>, prevState: Readonly<any>, snapshot?: any): void {
    if (this.props.siteUrl !== prevProps.siteUrl ||
      this.props.listName !== prevProps.listName) {
      this._queryList().then().catch(reason => console.error(reason));
    }
  }

  private _onSearch = (newValue: string): void => {
    let newListItems;
    if (newValue) {
      newListItems = filterBy(this.state.allItems, {
        logic: "or",
        filters: [
          { field: this.props.questionFieldName, operator: "contains", value: newValue },
          { field: this.props.answerFieldName, operator: "contains", value: newValue },
          { field: this.props.subtitleFieldName, operator: "contains", value: newValue },
          { field: 'Topic', operator: "contains", value: newValue },
        ]
      });
    }
    else {
      newListItems = this.state.allItems;
    }

    this.setState({ items: newListItems });
  }

  public render(): React.ReactElement<IFaqAccordionProps> {
    if (this.state.items === undefined) {
      return <div>Loading...</div>;
    }
    else {
      return (
        <div>
          <h2>{this.props.description}</h2>
          <SearchBox placeholder={`Search by Question, Answer, ${this.props.subtitleFieldName}, or Topic.`} onChange={(event, newValue) => this._onSearch(newValue)} />
          {this.state.items.map((item: any, index: number) => (
            <ExpansionPanel
              title={item[this.props.questionFieldName]}
              // subtitle={this.props.subtitleFieldName && <div style={{ textAlign: 'right' }}>{item[this.props.subtitleFieldName]}</div>}
              subtitle={
                this.props.subtitleFieldName &&
                typeof item[this.props.subtitleFieldName] === "string" &&
                <div style={{ textAlign: 'right' }}>{item[this.props.subtitleFieldName]}</div>
              }
              expanded={this.state.expanded === item.ID}
              tabIndex={0}
              key={item.ID}
              onAction={(event: ExpansionPanelActionEvent) => {
                this.setState({ expanded: event.expanded ? "" : item.ID });
              }}
            >
              <Reveal>
                {this.state.expanded === item.ID && (
                  <ExpansionPanelContent>
                    <div className="content">
                      <span className="content-text">
                        {
                          this.state.canUserEditListItems &&
                          <div>
                            <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
                              You have permissions to edit this list item.
                              <Link href={`${this.props.siteUrl}/Lists/${this.props.listName}/EditForm.aspx?ID=${item.ID}`} target="_blank" underline>
                                Click Here to Edit Item.
                              </Link>
                            </MessageBar>
                          </div>
                        }
                        <RichText value={item[this.props.answerFieldName]} isEditMode={false} />
                      </span>
                    </div>
                  </ExpansionPanelContent>
                )}
              </Reveal>
            </ExpansionPanel>
          ))}
        </div>
      );
    }
  }
}
