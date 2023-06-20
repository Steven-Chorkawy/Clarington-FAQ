import * as React from 'react';
import { IFaqAccordionProps } from './IFaqAccordionProps';
import { getSiteSP } from '../../../pnpjs-config';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import "@pnp/sp/items/get-all";
import "@pnp/sp/taxonomy";
import { ExpansionPanel, ExpansionPanelActionEvent, ExpansionPanelContent } from '@progress/kendo-react-layout';
import { Reveal } from '@progress/kendo-react-animation';
import "@pnp/sp/taxonomy";
import { SearchBox } from 'office-ui-fabric-react';

const MOC_ORG_GROUP_ID = '4026b60c-6222-432f-b07d-89c2396e8e64';
const DEPARTMENT_TERM_SET_ID = '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f';

export default class FaqAccordion extends React.Component<IFaqAccordionProps, any> {

  constructor(props: any) {
    super(props);
    console.log('ctor');
    console.log(props);
    this.state = {
      items: undefined
    };
    this._queryList();
  }

  // TODO: This method should query any managed metadata field not just a department field.
  private async _queryDepartmentName(termID: string): Promise<void> {
    if (!termID)
      return;

    // check if state has already been set. 
    if (this.state[termID]) {
      console.log(`${termID} has already been found: ${this.state[termID]}`);
    }
    else {
      console.log(`${termID} NOT FOUND!`);
      let res = await getSiteSP().termStore.groups.getById(MOC_ORG_GROUP_ID).sets.getById(DEPARTMENT_TERM_SET_ID).getTermById(termID)();

      if (res.labels) {
        if (res.labels.length > 0) {
          // There should only ever be one item in this array so just get the first one.
          this.setState({ [termID]: res.labels[0].name });

          // Replace the Department object with a simple string of the department name.  
          // This will make searching and rendering this info much easier.
          this.state.items.map((item: any, index: number) => {
            if (item.Department?.TermGuid === termID) {
              item.DepartmentDate = { ...item.Department }
              item.Department = res.labels[0].name
            }
          });
        }
      }
    }
  }

  private async _queryList(): Promise<void> {
    let listItems = await getSiteSP().web.lists.getByTitle(this.props.listName).items.getAll();

    listItems.forEach(item => {
      this._queryDepartmentName(item.Department?.TermGuid);
    });

    // Sort by Created date.  Newest to Oldest.
    let sortedList = listItems.sort((p1, p2) => (p1.Created < p2.Created) ? 1 : (p1.Created > p2.Created) ? -1 : 0);

    this.setState({ items: sortedList });
  }

  componentDidUpdate(prevProps: Readonly<IFaqAccordionProps>, prevState: Readonly<any>, snapshot?: any): void {
    if (this.props.siteUrl !== prevProps.siteUrl ||
      this.props.listName !== prevProps.listName) {
      this._queryList();
    }
  }

  public render(): React.ReactElement<IFaqAccordionProps> {
    if (this.state.items === undefined) {
      return <div>Loading...</div>;
    }
    else {
      return (
        <div>
          <h2>{this.props.description}</h2>
          <SearchBox placeholder="This Search Box Does Not Work Yet..." onSearch={newValue => console.log('value is ' + newValue)} />
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
