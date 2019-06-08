//rename file app.js to app.tsx and copy this file to src folder//

import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, { HeroListItem } from './HeroList';
import Progress from './Progress';

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: 'Ribbon',
          primaryText: 'Achieve more with Office integration'
        },
        {
          icon: 'Unlock',
          primaryText: 'Unlock features and functionality'
        },
        {
          icon: 'Design',
          primaryText: 'Create and visualize like a pro'
        }
      ]
    });
  }

  click = async () => {
    try {
      await Excel.run(async context => {
       
        const range = context.workbook.getSelectedRange();
        
        range.load("address");

        
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }
  createTable = async () => {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
        expensesTable.name = "ExpensesTable"
        
        expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]]
        expensesTable.rows.add(null /*add at the end*/, [
          ["1/1/2017", "The Phone Company", "Communications", "120"],
          ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
          ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
          ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
          ["1/11/2017", "Bellows College", "Education", "350.1"],
          ["1/15/2017", "Trey Research", "Other", "135"],
          ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
          ]);
        
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();  

        await context.sync();        
      });
    } catch (error) {
      console.error(error);
    }
  }

  filterTable = async () => {
    try {
      await Excel.run(async context => {
       
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
        const categoryFilter = expensesTable.columns.getItem('Category').filter;
        categoryFilter.applyValuesFilter(["Education", "Groceries"]);
       
        
        await context.sync();        
      });
    } catch (error) {
      console.error(error);
    }
  }
  sortTable = async () => {
    try {
      await Excel.run(async context => {
      
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
        const sortFields = [
          {
              key: 1,            // Merchant column
              ascending: false,
          }
        ];

        expensesTable.sort.apply(sortFields);
             
        await context.sync();        
      });
    } catch (error) {
      console.error(error);
    }
  }
  createChart = async () => {
    try {
      await Excel.run(async context => {
       
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
        const dataRange = expensesTable.getDataBodyRange();
        
        const chart = currentWorksheet.charts.add('ColumnClustered', dataRange, "Auto");
        
        chart.setPosition("A15", "F30");
        chart.title.text = "Expenses";
        chart.legend.position = "Right"
        chart.legend.format.fill.setSolidColor("white");
        chart.dataLabels.format.font.size = 15;
        chart.dataLabels.format.font.color = "black";
        chart.series.getItemAt(0).name = 'Value in €';
                
        await context.sync();        
      });
    } catch (error) {
      console.error(error);
    }
  }
  freezeHeader = async () => {
    try {
      await Excel.run(async context => {
       
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        currentWorksheet.freezePanes.freezeRows(1);
        
        await context.sync();        
      });
    } catch (error) {
      console.error(error);
    }
  }

  render() {
    const {
      title,
      isOfficeInitialized,
    } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo='assets/logo-filled.png'
          message='Please sideload your addin to see app body.'
        />
      );
    }

    return (
      <div className='ms-welcome'>
        <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome' />
        <HeroList message='Discover what Office Add-ins can do for you today!' items={this.state.listItems}>
          <p className='ms-font-l'>Modify the source files, then click <b>Run</b>.</p>
          <Button className='ms-welcome__action' buttonType={ButtonType.command} iconProps={{ iconName: 'ChevronRight' }} onClick={this.createTable}>CreateTable</Button>
          <Button className='ms-welcome__action' buttonType={ButtonType.command} iconProps={{ iconName: 'ChevronRight' }} onClick={this.filterTable}>FilterTable</Button>
          <Button className='ms-welcome__action' buttonType={ButtonType.command} iconProps={{ iconName: 'ChevronRight' }} onClick={this.sortTable}>SortTable</Button>
          <Button className='ms-welcome__action' buttonType={ButtonType.command} iconProps={{ iconName: 'ChevronRight' }} onClick={this.createChart}>CreateChart</Button>          
          <Button className='ms-welcome__action' buttonType={ButtonType.command} iconProps={{ iconName: 'ChevronRight' }} onClick={this.freezeHeader}>FreezeHeader</Button>
          <Button className='ms-welcome__action' buttonType={ButtonType.primary} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>HighLight-yellow</Button>        
        </HeroList>
      </div>
    );
  }
}
