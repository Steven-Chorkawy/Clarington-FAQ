import * as React from 'react';
import { IFaqAccordionProps } from './IFaqAccordionProps';
import { Accordion } from '@pnp/spfx-controls-react';

export default class FaqAccordion extends React.Component<IFaqAccordionProps, {}> {
  public render(): React.ReactElement<IFaqAccordionProps> {
    let sampleItems: any = [
      { Question: "Q1", Response: "R1" },
      { Question: "Q2", Response: "R2" },
      { Question: "Q3", Response: "R3" },
      { Question: "Q4", Response: "R4" },
      { Question: "Q5", Response: "R5" },
      { Question: "Q6", Response: "R6" },
    ];
    return (
      <div>
        <p>{this.props.siteUrl} and {this.props.listName}</p>
        {sampleItems.map((item: any, index: number) => (
          <Accordion title={item.Question} defaultCollapsed={true} className={"itemCell"} key={index}>
            <div className={"itemContent"}>
              <div className={"itemResponse"}>{item.Response}</div>
            </div>
          </Accordion>
        ))}
      </div>
    );
  }
}
