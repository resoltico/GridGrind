package dev.erst.gridgrind.engine.customxml;

import dev.erst.gridgrind.core.model.value.ExcelDateSystem;
import dev.erst.gridgrind.core.model.value.TemporalSemantics;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.FieldValidationTemporalRule;
import dev.erst.gridgrind.core.problem.ValueInvariantException;
import java.time.LocalDate;
import javax.xml.datatype.DatatypeConstants;
import javax.xml.datatype.DatatypeFactory;
import javax.xml.datatype.XMLGregorianCalendar;
import org.apache.poi.xssf.extractor.XSSFImportFromXML.DateInput;

/** XML Schema lexical parsing at the adapter; date policy and serial meaning stay in core. */
final class MappedDates {
  private MappedDates() {}

  static double serial(DateInput input, ExcelDateSystem system) {
    XMLGregorianCalendar parsed = parse(input, system);
    if (!DatatypeConstants.DATE.equals(parsed.getXMLSchemaType()))
      throw refusal(input, system, problem(FieldValidationTemporalRule.XML_DATE_LEXICAL));
    if (parsed.getTimezone() != DatatypeConstants.FIELD_UNDEFINED)
      throw refusal(input, system, problem(FieldValidationTemporalRule.XML_DATE_OFFSET));
    if (parsed
                .getEonAndYear()
                .compareTo(
                    java.math.BigInteger.valueOf(
                        ExcelDateSystem.WINDOWS_1900.minimumDate().getYear()))
            < 0
        || parsed
                .getEonAndYear()
                .compareTo(java.math.BigInteger.valueOf(ExcelDateSystem.MAXIMUM_DATE.getYear()))
            > 0)
      throw refusal(
          input,
          system,
          problem(
              FieldValidationTemporalRule.DATE_RANGE,
              ExcelDateSystem.WINDOWS_1900.minimumDate().toString(),
              ExcelDateSystem.MAXIMUM_DATE.toString()));
    try {
      return TemporalSemantics.serial(
          LocalDate.of(parsed.getYear(), parsed.getMonth(), parsed.getDay()), system);
    } catch (ValueInvariantException invalid) {
      throw refusal(input, system, invalid);
    }
  }

  private static XMLGregorianCalendar parse(DateInput input, ExcelDateSystem system) {
    try {
      return DatatypeFactory.newDefaultInstance()
          .newXMLGregorianCalendar(input.lexicalValue().trim());
    } catch (IllegalArgumentException invalid) {
      var problem = problem(FieldValidationTemporalRule.XML_DATE_LEXICAL);
      problem.initCause(invalid);
      throw refusal(input, system, problem);
    }
  }

  private static MappedDateInputException refusal(
      DateInput input, ExcelDateSystem system, ValueInvariantException problem) {
    return new MappedDateInputException(problem, input.xpath(), input.mappingRowIndex(), system);
  }

  private static ValueInvariantException problem(
      FieldValidationTemporalRule rule, String... operands) {
    return new ValueInvariantException(
        FieldValidationProblem.detached("mapped DATE", rule, operands));
  }
}
