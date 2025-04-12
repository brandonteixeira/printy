export class Email {
  constructor(item) {
    debugger;
    this.item = item;
    this.htmlContent = this.#toHtml();
  }

  toPdf() {
    console.log(this.item + "in pdf");
  }

  // Method to display the person's information
  greet() {
    console.log(`Hello, my name is ${this.name} and I am ${this.age} years old.`);
  }

  // private methods
  async #toHtml() {
    return await new Promise((resolve) => {
      this.item.body.getAsync(Office.CoercionType.Html, (result) => {
        resolve(result.value);
      });
    });
  }
}
