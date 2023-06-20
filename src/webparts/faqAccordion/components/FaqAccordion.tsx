import * as React from 'react';
import { IFaqAccordionProps } from './IFaqAccordionProps';
import { getSiteSP } from '../../../pnpjs-config';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import "@pnp/sp/items/get-all";
import "@pnp/sp/taxonomy";
import { ExpansionPanel, ExpansionPanelActionEvent, ExpansionPanelContent } from '@progress/kendo-react-layout';
import { Reveal } from '@progress/kendo-react-animation';
import "@pnp/sp/taxonomy";

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

          this.state.items.map((item: any, index: number) => {
            if (item.Department?.TermGuid === termID) {
              item.Department = { ...item.Department, Name: res.labels[0].name }
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
          {this.state.items.map((item: any, index: number) => (
            <ExpansionPanel
              title={item[this.props.questionFieldName]}
              subtitle={<div style={{ textAlign: 'right' }}>{item.Department?.Name}</div>}
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
