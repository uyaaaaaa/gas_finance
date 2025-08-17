class Trigger {
  /**
   * Constructor.
   * @param {string} name
   */
  constructor(name) {
    this.name = name;
    this.executeAt = new Date();
    this.setDefault();
  }

  /**
   * Set defaullt trigger date.
   */
  setDefault() {
    this.executeAt.setDate(this.executeAt.getDate() + 1);
    this.executeAt.setHours(this.executeAt.getHours());
    this.executeAt.setMinutes(this.executeAt.getMinutes());
  }

  /**
   * @param {number} month
   * @return {Trigger}
   */
  setMonth(month) {
    this.executeAt.setMonth(month - 1);
    return this;
  }

  /**
   * @param {number} date
   * @return {Trigger}
   */
  setDate(date) {
    this.executeAt.setDate(date);
    return this;
  }

  /**
   * @param {number} hours
   * @return {Trigger}
   */
  setHours(hours) {
    this.executeAt.setHours(hours);
    return this;
  }

  /**
   * @param {number} minutes
   * @return {Trigger}
   */
  setMinutes(minutes) {
    this.executeAt.setMinutes(minutes);
    return this;
  }

  /**
   * Remove existing and create new.
   */
  update() {
    this.remove();
    this.create();
  }
  
  /**
   * Create new trigger (no delete).
   */
  create() {
    console.log(Utilities.formatDate(this.executeAt, "JST", "yyyy/MM/dd HH:mm"));

    ScriptApp.newTrigger(this.name)
      .timeBased()
      .at(this.executeAt)
      .create();
  }

  /**
   * Remove trigger by name if exists.
   */
  remove() {
    const existing = this.getExistingByName(this.name);

    if (existing === null) {
      return;
    }
    
    ScriptApp.deleteTrigger(existing);
  }

  /**
   * Get existing trigger by name.
   * @param {string} name
   * @return {null|ScriptApp.Trigger}
   */
  getExistingByName(name) {
    const triggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === name) {
        return triggers[i];
      }
    }

    return null;
  }
}
