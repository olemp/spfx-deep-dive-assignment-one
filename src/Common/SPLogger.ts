import { ServiceScope, Log } from '@microsoft/sp-core-library';

/**
 * Generic Logger
 * use CTRL + F12 to see the Developer Dashboard
 * Visit https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-developer-dashboard for more details on this
 */
export class SPLogger {

    private serviceScope: ServiceScope;
    constructor(scope: ServiceScope) {
        this.serviceScope = scope;
    }
    /**
    * Returns the caller method name, assuming it is third in error stack
    */
    private get getMethodName(): string {
        try {
            let stackTrace = (new Error()).stack; // Only tested in latest FF and Chrome
            let stack: any = stackTrace.replace(/^Error\s+/, ''); // Sanitize Chrome
            let callerName: string = stack.split(`\n`)[2]; // 1st item is this, 3nd item is original caller
            callerName = callerName.replace(/^\s+at Object./, ''); // Sanitize Chrome
            callerName = callerName.replace(/ \(.+\)$/, ''); // Sanitize Chrome
            callerName = callerName.replace(/\@.+/, '');
            callerName = callerName.replace(`prototype.`, ``);
            callerName = callerName.split(`at `).length > 1 ? callerName.split(`at `)[1] : callerName;
            return callerName.trim();
        } catch (error) {
            console.error(`Error in SPLogger.getMethodName: ${JSON.stringify(error)}`);
            return 'FoundNoMethodName';
        }
    }

    /**
     * Logs error to the console
     * @param message : Error object
     */
    public logError(message: Error): void {
        Log.error(this.getMethodName, message, this.serviceScope);
        console.log(this.getMethodName);
        console.error(message);
    }

    /**
     * Logs a general informational message.
     * @param message the message to be logged
     */
    public logInfo(message: any): void {
        Log.info(this.getMethodName, message, this.serviceScope);
        console.log(this.getMethodName);
        console.info(message);
    }

    /**
     * Logs a message which contains detailed information that is generally only needed for troubleshooting.
     * @param message the message to be logged
     */
    public logVerbose(message: any): void {
        Log.verbose(this.getMethodName, message, this.serviceScope);
        console.log(this.getMethodName);
        console.info(message);
    }

    /**
     * Logs a warning.
     * @param message the message to be logged
     */
    public logWarning(message: any): void {
        Log.warn(this.getMethodName, message, this.serviceScope);
        console.log(this.getMethodName);
        console.warn(message);
    }
}