import * as React from 'react';
import { IFaqAccordionProps } from './IFaqAccordionProps';
import { Accordion } from '@pnp/spfx-controls-react';
import { getSiteSP } from '../../../pnpjs-config';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";


export default class FaqAccordion extends React.Component<IFaqAccordionProps, any> {

  constructor(props: any) {
    super(props);
    console.log('ctor');
    console.log(props);
    this.state = {
      items: undefined
    };
    this._queryAndSetListState();
  }

  /**
   * Query a given SharePoint list to populate the accordion webpart.
   * @returns List items
   */
  private async _queryList(): Promise<any> {
    return await getSiteSP().web.lists.getByTitle(this.props.listName).items();
  }

  /**
   * Call the main query method and set the results in the state.
   */
  private _queryAndSetListState(): void {
    this._queryList()
      .then(value => {
        this.setState({ items: value });
      })
      .catch(value => {
        console.log(value);
        this.setState({ items: [] });
        alert('Failed to load list for Q&A webpart!');
      });
  }

  componentDidUpdate(prevProps: Readonly<IFaqAccordionProps>, prevState: Readonly<any>, snapshot?: any): void {
    if (this.props.siteUrl !== prevProps.siteUrl ||
      this.props.listName !== prevProps.listName) {
      this._queryAndSetListState();
    }
  }

  public render(): React.ReactElement<IFaqAccordionProps> {
    if (this.state.items === undefined) {
      return <div>Loading...</div>;
    }
    else {
      return (
        <div>
          <h2>{this.props.webPartTitle}</h2>
          {this.state.items.map((item: any, index: number) => (
            <Accordion title={item[this.props.questionFieldName]} defaultCollapsed={true} className={"itemCell"} key={index}>
              <div className={"itemContent"}>
                <div className={"itemResponse"}>
                  <RichText value={item[this.props.answerFieldName]} isEditMode={false} />
                </div>
              </div>
            </Accordion>
          ))}
        </div>
      );
    }
  }
}
