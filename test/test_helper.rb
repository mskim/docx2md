#!/usr/bin/env ruby
# encoding: utf-8

require 'minitest/autorun'
$LOAD_PATH.unshift File.join(File.dirname(__FILE__), '..', 'lib')
# $LOAD_PATH.unshift File.join(File.dirname(__FILE__), '..')
# $LOAD_PATH.unshift File.dirname(__FILE__), '../..'
require "docx2md"
include Docx2md
