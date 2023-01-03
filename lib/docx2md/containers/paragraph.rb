require 'docx2md/containers/text_run'
require 'docx2md/containers/container'

module Docx2md
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element
        attr_reader :paragraph_footnotes, :footnote_number
        attr_reader :text_with_footnote
        def self.tag
          'p'
        end


        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node, document_properties = {})
          @node = node
          @properties_tag = 'pPr'
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
          @paragraph_footnotes = []
        end

        # Set text of paragraph
        def text=(content)
          if text_runs.size == 1
            text_runs.first.text = content
          elsif text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            text_runs.each {|r| r.node.remove }
            new_r = TextRun.create_within(self)
            new_r.text = content
          end
        end

        # Return text of paragraph
        def to_s
          text_runs.map(&:text).join('')
        end

        # Return paragraph as a <p></p> HTML fragment with formatting based on properties.
        def to_html
          html = ''
          text_runs.each do |text_run|
            html << text_run.to_html
          end
          styles = { 'font-size' => "#{font_size}pt" }
          styles['text-align'] = alignment if alignment
          html_tag(:p, content: html, styles: styles)
        end

        def to_txt(document)
          footnotes_hash  = document.footnotes_hash
          footnote_number  = document.footnote_number
          text = ''
          text_runs.each do |text_run|
            md = text_run.to_markdown
            if md=~/\[\^(\d+?)\]/
              footnote_id = $1
              md = "[^#{footnote_number}]"
              footnote_descrption_text =  footnotes_hash[footnote_id]
              footnote_descrption  = "[^#{footnote_number}]: #{footnote_descrption_text}"
              footnote_number += 1
              @paragraph_footnotes << footnote_descrption
              text += md
            else
              text += md
            end
          end
          text
        end

        def to_markdown(document)
          styles_hash  = document.styles_hash
          footnotes_hash  = document.footnotes_hash
          footnote_number  = document.footnote_number
          text = ''
          text_runs.each do |text_run|
            md = text_run.to_markdown
            if md=~/\[\^(\d+?)\]/
              footnote_id = $1
              md = "[^#{footnote_number}]"
              footnote_descrption_text =  footnotes_hash[footnote_id]
              footnote_descrption  = "[^#{footnote_number}]: #{footnote_descrption_text}"
              footnote_number += 1
              @paragraph_footnotes << footnote_descrption
              text += md
            else
              text += md
            end
          end
          para_style_node =  @node.xpath('w:pPr/w:pStyle')
          style_id = para_style_node.attribute('val').value
          style_name = styles_hash[style_id]
          markdown_tag(paragraph_style_name,  text)
        end

        # we use `` mark to indicate blockquote
        # so, convert it to blockquote
        def convert_to_markdown(text)
          if text=~/``/
            content = text.split("\n").map do |line|
              "> #{line}"
            end
            content = content.join("\n")
          else
            text += "\n"
          end
        end

        def markdown_tag(paragraph_style_name,  text)
          markup = ""
          content = text
          case paragraph_style_name
          when "Heading 1", "장제목"
            markup = "# "
          when "Heading 2", /좌즉/
            markup = "## "
          when "Heading 3", /우측/
            markup = "### "
            when "Heading 4"
            markup = "#### "
          when "Heading 5"
            markup = "##### "
          when "Heading 6"
            markup = "###### "
          when "Block Text", /인용/
            markup = ""
            content = text.split("\n").map do |line|
              "> #{line}"
            end
            content = content.join("\n")
          else
            markup = ""
          end
          para_text = markup + content + "\n"
          @paragraph_footnotes.each_with_index do |note, i|
            para_text += "\n"
            para_text += "#{note}"
            para_text += "\n"
          end
          para_text
        end
        

        # Array of text runs contained within paragraph
        def text_runs
          @node.xpath('w:r|w:hyperlink').map { |r_node| Containers::TextRun.new(r_node, @document_properties) }
        end

        # Iterate over each text run within a paragraph
        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end

        def aligned_left?
          ['left', nil].include?(alignment)
        end

        def aligned_right?
          alignment == 'right'
        end

        def aligned_center?
          alignment == 'center'
        end

        def font_size
          size_tag = @node.xpath('w:pPr//w:sz').first
          size_tag ? size_tag.attributes['val'].value.to_i / 2 : @font_size
        end
        
        alias_method :text, :to_s

        private

        # Returns the alignment if any, or nil if left
        def alignment
          alignment_tag = @node.xpath('.//w:jc').first
          alignment_tag ? alignment_tag.attributes['val'].value : nil
        end

      end
    end
  end
end
