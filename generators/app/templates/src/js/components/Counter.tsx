import * as React from "react";

export interface CounterProps { demoSettings: string; }
export interface CounterState { currentCount: number; }

export class Counter extends React.Component<CounterProps, CounterState> {
    constructor() {
        super();
        this.state = { currentCount: 0 };
    }

    render() {
        return <div>
            <h3>Hello from {this.props.demoSettings}!</h3>
            <p>This is a simple example of a React component.</p>
            <p>Current count: <strong>{ this.state.currentCount }</strong></p>
            <button type='button' onClick={ () => this.incrementCounter() }>Increment</button>
        </div>;
    }

    incrementCounter() {
        this.setState({
            currentCount: this.state.currentCount + 1
        })
    }
}