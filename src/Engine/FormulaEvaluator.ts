import Cell from "./Cell";
import SheetMemory from "./SheetMemory";
import { ErrorMessages } from "./GlobalDefinitions";
import { stat } from "fs";

export class FormulaEvaluator {
  // Define a function called update that takes a string parameter and returns a number
  private _errorOccured: boolean = false;
  private _errorMessage: string = "";
  private _currentFormula: FormulaType = [];
  private _lastResult: number = 0;
  private _sheetMemory: SheetMemory;
  private _result: number = 0;

  constructor(memory: SheetMemory) {
    this._sheetMemory = memory;
  }

  validate(formula: FormulaType): FormulaType {
    let lastValidIndex = 0;
    if (formula.length === 0) {
      this._errorMessage = ErrorMessages.emptyFormula;
      return [];
    }

    let state = "initial";
    let brackets = 0;
    for (const [i, element] of formula.entries()) {
      if (element === "(") {
        brackets += 1;
      }
      if (element === ")") {
        brackets -= 1;
      }
      if (brackets < 0) {
        this._errorMessage = ErrorMessages.missingParentheses;
        return formula.slice(0, lastValidIndex + 1);
      }

      switch (state) {
        case "initial":
          if (this.isNumber(element) || this.isCellReference(element)) {
            state = "value";
            if (brackets === 0) {
              lastValidIndex = i;
            }
          } else if (element === "-") {
            state = "minus";
          } else if (element === "(") {
            state = "initial";
          } else {
            this._errorMessage = ErrorMessages.invalidOperator;
            return formula.slice(0, lastValidIndex + 1);
          }
          break;
        case "minus":
          if (
            this.isNumber(element) ||
            this.isCellReference(element) ||
            element === "("
          ) {
            state = "value";
            if (brackets === 0) {
              lastValidIndex = i;
            }
          } else {
            this._errorMessage = ErrorMessages.invalidOperator;
            return formula.slice(0, lastValidIndex + 1);
          }
          break;
        case "value":
          if (
            this.isNumber(element) ||
            this.isCellReference(element) ||
            element === "("
          ) {
            this._errorMessage = ErrorMessages.invalidOperator;
            return formula.slice(0, lastValidIndex + 1);
          } else if (element === ")") {
            state = "value";
            if (brackets === 0) {
              lastValidIndex = i;
            }
          } else {
            state = "operator";
          }

          break;
        case "operator":
          if (this.isNumber(element) || this.isCellReference(element)) {
            state = "value";
            if (brackets === 0) {
              lastValidIndex = i;
            }
          } else if (element === "(") {
            state = "initial";
          } else {
            this._errorMessage = ErrorMessages.invalidOperator;
            return formula.slice(0, lastValidIndex + 1);
          }
      }
    }

    if (brackets !== 0) {
      this._errorMessage = ErrorMessages.missingParentheses;
      return formula.slice(0, lastValidIndex + 1);
    }

    if (state === "operator" || state === "minus") {
      this._errorMessage = ErrorMessages.partial;
      return formula.slice(0, lastValidIndex + 1);
    }

    return formula.slice(0, lastValidIndex + 1);
  }

  evaluate(formula: FormulaType) {
    this._errorMessage = "";
    const validatedFormula = this.validate(formula);

    this._result = this.getValue([...validatedFormula]);
  }

  evaluateBrackets(formula: FormulaType) {
    let brackets: number = 0;
    let bracketStart: number = -1;
    let bracketEnd: number = -1;

    for (let i = formula.length - 1; i >= 0; i--) {
      const element = formula[i];
      if (element === ")") {
        brackets += 1;
        if (brackets === 1) {
          bracketEnd = i;
        }
      } else if (element === "(") {
        brackets -= 1;
        if (brackets === 0) {
          bracketStart = i;
          let value = this.getValue(
            formula.slice(bracketStart + 1, bracketEnd)
          );
          if (this._errorMessage !== "") {
            return this._result;
          }
          formula.splice(bracketStart, bracketEnd - bracketStart + 1, value); // Replace from bracketStart to bracketEnd
        }
      }
    }
  }

  evaluateCellReference(formula: FormulaType) {
    for (let i = formula.length - 1; i >= 0; i--) {
      const element = formula[i];
      if (this.isCellReference(element)) {
        let [value, error] = this.getCellValue(element);
        if (error !== "") {
          this._errorMessage = error;
        }
        formula[i] = value;
      }
    }
  }

  evaluateMultiplicationAndDivision(formula: FormulaType) {
    for (let i = formula.length - 1; i >= 0; i--) {
      const element = formula[i];
      if (element == "/") {
        let right = Number(formula[i + 1]);
        if (right === 0) {
          this._result = Infinity;
          this._errorMessage = ErrorMessages.divideByZero;
        }
        formula.splice(i, 1, "*");
        formula.splice(i + 1, 1, 1 / right);
      }
    }
    for (let i = formula.length - 1; i >= 0; i--) {
      const element = formula[i];
      if (element === "*") {
        let left = Number(formula[i - 1]);
        let right = Number(formula[i + 1]);
        let result = left * right;
        formula.splice(i - 1, 3, result);
      }
    }
  }

  evaluateAdditionAndSubtraction(formula: FormulaType) {
    for (let i = formula.length - 1; i >= 0; i--) {
      const element = formula[i];
      if (element === "-") {
        let right = Number(formula[i + 1]);
        if (i - 1 < 0 || !this.isNumber(formula[i - 1])) {
          formula.splice(i, 2, -right);
        } else {
          formula.splice(i, 1, "+");
          formula.splice(i + 1, 1, -right);
        }
      }
    }
    for (let i = formula.length - 1; i >= 0; i--) {
      const element = formula[i];
      if (element == "+") {
        let left = Number(formula[i - 1]);
        let right = Number(formula[i + 1]);
        let result = left + right;
        formula.splice(i - 1, 3, result);
      }
    }
  }

  getValue(formula: FormulaType): number {
    this.evaluateBrackets(formula);

    this.evaluateCellReference(formula);

    this.evaluateMultiplicationAndDivision(formula);

    this.evaluateAdditionAndSubtraction(formula);

    if (formula.length > 1) {
      this._errorMessage = ErrorMessages.invalidFormula;
      return this._result;
    }

    return this.isNumber(formula[0]) ? Number(formula[0]) : 0;
  }

  public get error(): string {
    return this._errorMessage;
  }

  public get result(): number {
    return this._result;
  }

  /**
   *
   * @param token
   * @returns true if the toke can be parsed to a number
   */
  isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

  /**
   *
   * @param token
   * @returns true if the token is a cell reference
   *
   */
  isCellReference(token: TokenType): boolean {
    return Cell.isValidCellLabel(token);
  }

  /**
   *
   * @param token
   * @returns [value, ""] if the cell formula is not empty and has no error
   * @returns [0, error] if the cell has an error
   * @returns [0, ErrorMessages.invalidCell] if the cell formula is empty
   *
   */
  getCellValue(token: TokenType): [number, string] {
    let cell = this._sheetMemory.getCellByLabel(token);
    let formula = cell.getFormula();
    let error = cell.getError();

    // if the cell has an error return 0
    if (error !== "" && error !== ErrorMessages.emptyFormula) {
      return [0, error];
    }

    // if the cell formula is empty return 0
    if (formula.length === 0) {
      return [0, ErrorMessages.invalidCell];
    }

    let value = cell.getValue();
    return [value, ""];
  }
}

export default FormulaEvaluator;
