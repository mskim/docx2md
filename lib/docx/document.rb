require 'docx2md/containers'
require 'docx2md/elements'
require 'nokogiri'
require 'zip'

module Docx2md
  # The Document class wraps around a docx file and provides methods to
  # interface with it.
  #
  #   # get a Docx::Document for a docx file in the local directory
  #   doc = Docx::Document.open("test.docx")
  #
  #   # get the text from the document
  #   puts doc.text
  #
  #   # do the same thing in a block
  #   Docx::Document.open("test.docx") do |d|
  #     puts d.text
  #   end
  class Document
    attr_reader :xml, :doc, :zip, :styles
    attr_reader :styles_hash, :footnotes_hash

    def initialize(path_or_io, options = {})
      @replace = {}

      # if path-or_io is string && does not contain a null byte
      if (path_or_io.instance_of?(String) && !/\u0000/.match?(path_or_io))
        @zip = Zip::File.open(path_or_io)
      else
        @zip = Zip::File.open_buffer(path_or_io)
      end

      document = @zip.glob('word/document*.xml').first
      raise Errno::ENOENT if document.nil?

      @document_xml = document.get_input_stream.read
      @doc = Nokogiri::XML(@document_xml)
      load_styles
      yield(self) if block_given?
    ensure
      @zip.close
    end

    # This stores the current global document properties, for now
    def document_properties
      {
        font_size: font_size,
        hyperlinks: hyperlinks
      }
    end

    # With no associated block, Docx::Document.open is a synonym for Docx::Document.new. If the optional code block is given, it will be passed the opened +docx+ file as an argument and the Docx::Document oject will automatically be closed when the block terminates. The values of the block will be returned from Docx::Document.open.
    # call-seq:
    #   open(filepath) => file
    #   open(filepath) {|file| block } => obj
    def self.open(path, &block)
      new(path, &block)
    end

    def paragraphs
      @doc.xpath('//w:document//w:body/w:p').map { |p_node| parse_paragraph_from p_node }
    end

    def bookmarks
      bkmrks_hsh = {}
      bkmrks_ary = @doc.xpath('//w:bookmarkStart').map { |b_node| parse_bookmark_from b_node }
      # auto-generated by office 2010
      bkmrks_ary.reject! { |b| b.name == '_GoBack' }
      bkmrks_ary.each { |b| bkmrks_hsh[b.name] = b }
      bkmrks_hsh
    end

    def to_xml
      Nokogiri::XML(@document_xml)
    end

    def tables
      @doc.xpath('//w:document//w:body//w:tbl').map { |t_node| parse_table_from t_node }
    end

    # Some documents have this set, others don't.
    # Values are returned as half-points, so to get points, that's why it's divided by 2.
    def font_size
      return nil unless @styles

      size_tag = @styles.xpath('//w:docDefaults//w:rPrDefault//w:rPr//w:sz').first
      size_tag ? size_tag.attributes['val'].value.to_i / 2 : nil
    end

    # Hyperlink targets are extracted from the document.xml.rels file
    def hyperlinks
      hyperlink_relationships.each_with_object({}) do |rel, hash|
        hash[rel.attributes['Id'].value] = rel.attributes['Target'].value
      end
    end

    def hyperlink_relationships
      @rels.xpath("//xmlns:Relationship[contains(@Type,'hyperlink')]")
    end

    ##
    # *Deprecated*
    #
    # Iterates over paragraphs within document
    # call-seq:
    #   each_paragraph => Enumerator
    def each_paragraph
      paragraphs.each { |p| yield(p) }
    end

    # call-seq:
    #   to_s -> string
    def to_s
      paragraphs.map(&:to_s).join("\n")
    end

    # Output entire document as a String HTML fragment
    def to_html
      paragraphs.map(&:to_html).join("\n")
    end

    def to_markdown
      build_styles_hash
      build_footnotes_hash
      paragraphs.map{|p| p.to_markdown(self)}.join("\n")
    end

    def build_styles_hash
      # get style_name from style.xml
      hash = {}
      style_list =styles.xpath('//w:style')
      style_list.each do |style|        
        style_id = style.attribute('styleId').value
        style_name_node  = style.children.first
        style_name = style_name_node.attribute('val').value
        hash[style_id]  = style_name
      end
      @styles_hash = hash
    end

    def build_footnotes_hash
      @footnote_xml = @zip.read('word/footnotes.xml')
      @footnotes = Nokogiri::XML(@footnote_xml)
      @footnotes_hash = {}
      footnotes.each do |footnote|
        binding.pry
        footnote_id = footnote.attribute('id').value
        footnote_text_node = footnote.at_xpath('.//w:t')
        footnote_text = footnote_text_node.text
        @footnotes_hash[footnote_id] = footnote_text
      end
      @footnotes_hash
    end

    # Save document to provided path
    # call-seq:
    #   save(filepath) => void
    def save(path)
      update
      Zip::OutputStream.open(path) do |out|
        zip.each do |entry|
          next unless entry.file?

          out.put_next_entry(entry.name)

          if @replace[entry.name]
            out.write(@replace[entry.name])
          else
            out.write(zip.read(entry.name))
          end
        end
      end
      zip.close
    end

    # Output entire document as a StringIO object
    def stream
      update
      stream = Zip::OutputStream.write_buffer do |out|
        zip.each do |entry|
          next unless entry.file?

          out.put_next_entry(entry.name)

          if @replace[entry.name]
            out.write(@replace[entry.name])
          else
            out.write(zip.read(entry.name))
          end
        end
      end

      stream.rewind
      stream
    end

    alias text to_s

    def replace_entry(entry_path, file_contents)
      @replace[entry_path] = file_contents
    end

    private

    def load_styles
      @styles_xml = @zip.read('word/styles.xml')
      @styles = Nokogiri::XML(@styles_xml)
      load_rels
    rescue Errno::ENOENT => e
      warn e.message
      nil
    end



    def load_rels
      rels_entry = @zip.glob('word/_rels/document*.xml.rels').first
      raise Errno::ENOENT unless rels_entry

      @rels_xml = rels_entry.get_input_stream.read
      @rels = Nokogiri::XML(@rels_xml)
    end

    #--
    # TODO: Flesh this out to be compatible with other files
    # TODO: Method to set flag on files that have been edited, probably by inserting something at the
    # end of methods that make edits?
    #++
    def update
      replace_entry 'word/document.xml', doc.serialize(save_with: 0)
    end

    # generate Elements::Containers::Paragraph from paragraph XML node
    def parse_paragraph_from(p_node)
      Elements::Containers::Paragraph.new(p_node, document_properties)
    end

    # generate Elements::Bookmark from bookmark XML node
    def parse_bookmark_from(b_node)
      Elements::Bookmark.new(b_node)
    end

    def parse_table_from(t_node)
      Elements::Containers::Table.new(t_node)
    end
  end
end
