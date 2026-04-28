package dev.erst.gridgrind.excel;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandleProxies;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.lang.invoke.VarHandle;
import java.util.Objects;

/** Centralizes reflective access to POI private members behind one engine-owned seam. */
final class PoiPrivateAccessSupport {
  private PoiPrivateAccessSupport() {}

  static VarHandle requireVarHandle(
      MethodHandles.Lookup lookup, PoiPrivateContract contract, Class<?> fieldType) {
    Objects.requireNonNull(contract, "contract must not be null");
    return requireVarHandle(
        lookup, contract.owner(), contract.lookupName(), fieldType, contract.failureMessage());
  }

  static VarHandle requireVarHandle(
      MethodHandles.Lookup lookup,
      Class<?> owner,
      String fieldName,
      Class<?> fieldType,
      String failureMessage) {
    Objects.requireNonNull(lookup, "lookup must not be null");
    Objects.requireNonNull(owner, "owner must not be null");
    Objects.requireNonNull(fieldName, "fieldName must not be null");
    Objects.requireNonNull(fieldType, "fieldType must not be null");
    Objects.requireNonNull(failureMessage, "failureMessage must not be null");
    try {
      return MethodHandles.privateLookupIn(owner, lookup)
          .findVarHandle(owner, fieldName, fieldType);
    } catch (ReflectiveOperationException exception) {
      throw new IllegalStateException(failureMessage, exception);
    }
  }

  static MethodHandle requireConstructor(
      MethodHandles.Lookup lookup, PoiPrivateContract contract, MethodType constructorType) {
    Objects.requireNonNull(contract, "contract must not be null");
    return requireConstructor(lookup, contract.owner(), constructorType, contract.failureMessage());
  }

  static MethodHandle requireConstructor(
      MethodHandles.Lookup lookup,
      Class<?> owner,
      MethodType constructorType,
      String failureMessage) {
    Objects.requireNonNull(lookup, "lookup must not be null");
    Objects.requireNonNull(owner, "owner must not be null");
    Objects.requireNonNull(constructorType, "constructorType must not be null");
    Objects.requireNonNull(failureMessage, "failureMessage must not be null");
    try {
      return MethodHandles.privateLookupIn(owner, lookup).findConstructor(owner, constructorType);
    } catch (ReflectiveOperationException exception) {
      throw new IllegalStateException(failureMessage, exception);
    }
  }

  static MethodHandle requireVirtual(
      MethodHandles.Lookup lookup, PoiPrivateContract contract, MethodType methodType) {
    Objects.requireNonNull(contract, "contract must not be null");
    return requireVirtual(
        lookup, contract.owner(), contract.lookupName(), methodType, contract.failureMessage());
  }

  static MethodHandle requireVirtual(
      MethodHandles.Lookup lookup,
      Class<?> owner,
      String methodName,
      MethodType methodType,
      String failureMessage) {
    Objects.requireNonNull(lookup, "lookup must not be null");
    Objects.requireNonNull(owner, "owner must not be null");
    Objects.requireNonNull(methodName, "methodName must not be null");
    Objects.requireNonNull(methodType, "methodType must not be null");
    Objects.requireNonNull(failureMessage, "failureMessage must not be null");
    try {
      return MethodHandles.privateLookupIn(owner, lookup)
          .findVirtual(owner, methodName, methodType);
    } catch (ReflectiveOperationException exception) {
      throw new IllegalStateException(failureMessage, exception);
    }
  }

  @SuppressWarnings("unchecked")
  static <T> T asInterfaceInstance(Class<T> interfaceType, MethodHandle handle) {
    Objects.requireNonNull(interfaceType, "interfaceType must not be null");
    Objects.requireNonNull(handle, "handle must not be null");
    return MethodHandleProxies.asInterfaceInstance(interfaceType, handle);
  }
}
