import * as React from "react";
import styles from "./WeeklyReport.module.scss";
import { IWeeklyReportProps } from "./IWeeklyReportProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp";
import { graph } from "@pnp/graph";
import MainComponent_WeeklyReport from "./MainComponent_WeeklyReport";

export default class WeeklyReport extends React.Component<
  IWeeklyReportProps,
  {}
> {
  constructor(prop: IWeeklyReportProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }

  public render(): React.ReactElement<IWeeklyReportProps> {
    const {} = this.props;

    return (
      <MainComponent_WeeklyReport
        context={sp}
        spcontext={this.props.context}
        graphContext={graph}
      />
    );
  }
}
