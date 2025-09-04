// import { Component } from '@angular/core';
// import { RouterOutlet } from '@angular/router';

// @Component({
//   selector: 'app-root',
//   standalone: true,
//   imports: [RouterOutlet],
//   templateUrl: './app.component.html',
//   styleUrl: './app.component.css'
// })
// export class AppComponent {
//   title = 'calculator-app';
// }

import { Component, NgZone } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-calculator',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class CalculatorComponent {
  display: string = '0';
  buttons: string[] = [
    '7',
    '8',
    '9',
    '/',
    '4',
    '5',
    '6',
    '*',
    '1',
    '2',
    '3',
    '-',
    '0',
    '.',
    '=',
    '+',
    'C',
  ];
  private firstNumber: number | null = null;
  private operator: string | null = null;
  private waitingForSecondNumber: boolean = false;

  constructor(private ngZone: NgZone) {}

  getButtonClass(button: string): string {
    if (button === '=') return 'btn btn-primary';
    if (button === 'C') return 'btn btn-danger';
    if (['+', '-', '*', '/'].includes(button)) return 'btn btn-warning';
    return 'btn btn-secondary';
  }

  onButtonClick(button: string): void {
    this.ngZone.run(() => {
      if (button === 'C') {
        this.clear();
      } else if (button === '=') {
        this.calculate();
      } else if (['+', '-', '*', '/'].includes(button)) {
        this.handleOperator(button);
      } else {
        this.handleNumber(button);
      }
    });
  }

  private clear(): void {
    this.display = '0';
    this.firstNumber = null;
    this.operator = null;
    this.waitingForSecondNumber = false;
  }

  private handleNumber(button: string): void {
    if (this.waitingForSecondNumber) {
      this.display = button;
      this.waitingForSecondNumber = false;
    } else {
      this.display = this.display === '0' ? button : this.display + button;
    }
  }

  private handleOperator(operator: string): void {
    if (this.firstNumber === null) {
      this.firstNumber = parseFloat(this.display);
      this.operator = operator;
      this.waitingForSecondNumber = true;
    } else if (!this.waitingForSecondNumber) {
      this.calculate();
      this.operator = operator;
      this.firstNumber = parseFloat(this.display);
      this.waitingForSecondNumber = true;
    }
  }

  private calculate(): void {
    if (
      this.firstNumber === null ||
      this.operator === null ||
      this.waitingForSecondNumber
    ) {
      return;
    }
    const secondNumber = parseFloat(this.display);
    let result: number;
    switch (this.operator) {
      case '+':
        result = this.firstNumber + secondNumber;
        break;
      case '-':
        result = this.firstNumber - secondNumber;
        break;
      case '*':
        result = this.firstNumber * secondNumber;
        break;
      case '/':
        result = this.firstNumber / secondNumber;
        break;
      default:
        return;
    }
    this.display = result.toString();
    this.firstNumber = null;
    this.operator = null;
    this.waitingForSecondNumber = false;
  }
}
