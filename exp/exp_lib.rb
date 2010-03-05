# Quick and dirty DDE Library

require 'ffi'
require 'win/dde'
require 'win/gui/message'
require_relative 'exp_lib'

module DDELib
  extend FFI::Library
  CP_WINANSI = 1004
  DNS_REGISTER = 1
  APPCLASS_STANDARD = 0
  CF_TEXT = 1

  XTYPF_NOBLOCK = 0x0002
  XCLASS_BOOL   = 0x1000
  XCLASS_FLAGS  = 0x4000
  XTYP_CONNECT  = 0x0060 | XCLASS_BOOL | XTYPF_NOBLOCK
  XTYP_POKE     = 0x0090 | XCLASS_FLAGS
  XTYP_EXECUTE  = 0x0050 | XCLASS_FLAGS
  TIMEOUT_ASYNC = 0xFFFFFFFF

  DDE_FACK      = 0x8000

  ffi_lib 'user32', 'kernel32'  # Default library
  ffi_convention :stdcall

  callback :DdeCallback, [:uint, :uint, :ulong, :pointer, :pointer, :pointer, :pointer], :ulong

  attach_function(:DdeInitializeA, [:pointer, :DdeCallback, :uint32, :uint32], :uint)
  attach_function(:DdeCreateStringHandleA, [:uint32, :pointer, :int], :ulong)
  attach_function :DdeNameService, [:uint32, :ulong, :ulong, :uint], :ulong
  attach_function(:DdeConnect, [:uint32, :ulong, :ulong, :pointer], :ulong)
  attach_function :DdeDisconnect, [:ulong], :int
  attach_function(:DdeClientTransaction, [:pointer, :uint32, :ulong, :ulong, :uint, :uint, :uint32, :pointer], :pointer)
  attach_function :DdeGetLastError, [:uint32], :int
end

