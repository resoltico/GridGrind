package dev.erst.gridgrind.jazzer.tool;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Describes the structured meaning of a replayed fuzz input. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ProtocolRequestDetails.class, name = "PROTOCOL_REQUEST"),
  @JsonSubTypes.Type(value = ProtocolWorkflowDetails.class, name = "PROTOCOL_WORKFLOW"),
  @JsonSubTypes.Type(value = CommandSequenceDetails.class, name = "ENGINE_COMMAND_SEQUENCE"),
  @JsonSubTypes.Type(value = XlsxRoundTripDetails.class, name = "XLSX_ROUND_TRIP")
})
public sealed interface ReplayDetails
    permits CommandSequenceDetails,
        ProtocolRequestDetails,
        ProtocolWorkflowDetails,
        XlsxRoundTripDetails {}
