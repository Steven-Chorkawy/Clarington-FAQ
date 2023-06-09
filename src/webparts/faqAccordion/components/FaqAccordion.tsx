import * as React from 'react';
import { IFaqAccordionProps } from './IFaqAccordionProps';
import { Accordion } from '@pnp/spfx-controls-react';
import { getSiteSP } from '../../../pnpjs-config';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import "@pnp/sp/items/get-all";
import "@pnp/sp/taxonomy";

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

  private async _queryList(): Promise<void> {
    getSiteSP().web.lists.getByTitle(this.props.listName).items.getAll().then(value => {
      this.setState({ items: value });
    });
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
