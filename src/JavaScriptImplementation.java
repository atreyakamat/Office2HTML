import java.io.PrintWriter;

public class JavaScriptImplementation {

    public void functionOne(PrintWriter writer) {
        writer.println("function functionOne() {");
        writer.println("    console.log('Hello from functionOne');");
        writer.println("}");
    }

    public void functionTwo(PrintWriter writer) {
        writer.println("function functionTwo() {");
        writer.println("    console.log('Hello from functionTwo');");
        writer.println("}");
    }

    public void functionThree(PrintWriter writer) {
        writer.println("function functionThree() {");
        writer.println("    console.log('Hello from functionThree');");
        writer.println("}");
    }

    public void functionFour(PrintWriter writer) {
        writer.println("function functionFour() {");
        writer.println("    console.log('Hello from functionFour');");
        writer.println("}");
    }

    public void functionFive(PrintWriter writer) {
        writer.println("function functionFive() {");
        writer.println("    console.log('Hello from functionFive');");
        writer.println("}");
    }
    public void writeScript(PrintWriter writer) {
        writer.println("<script>");
        functionOne(writer);
        functionTwo(writer);
        functionThree(writer);
        functionFour(writer);
        functionFive(writer);
        writer.println("</script>");
    }
}


