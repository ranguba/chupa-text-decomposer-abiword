# Copyright (C) 2019  Sutou Kouhei <kou@clear-code.com>
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA

require "tempfile"

module ChupaText
  module Decomposers
    class AbiWord < Decomposer
      include Loggable

      registry.register("abiword", self)

      EXTENSIONS = [
        "abw",
        "doc",
        "docx",
        "odt",
        "rtf",
        "zabw",
      ]
      MIME_TYPES = [
        "application/msword",
        "application/rtf",
        "application/vnd.oasis.opendocument.text",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/x-abiword",
      ]

      def initialize(options)
        super
        @command = find_command
        debug do
          if @command
            "#{log_tag}[command][found] #{@command.path}"
          else
            "#{log_tag}[command][not-found]"
          end
        end
      end

      def target?(data)
        return false if @command.nil?
        EXTENSIONS.include?(data.extension) or
          MIME_TYPES.include?(data.mime_type)
      end

      def target_score(data)
        if target?(data)
          10
        else
          nil
        end
      end

      def decompose(data)
        create_tempfiles(data) do |text, stdout, stderr|
          succeeded = @command.run("--to", "text",
                                   "--to-name", text.path,
                                   data.path.to_s,
                                   {
                                     data: data,
                                     spawn_options: {
                                       out: stdout.path,
                                       err: stderr.path,
                                     },
                                   })
          unless succeeded
            error do
              tag = "#{log_tag}[convert][exited][abnormally]"
              [
                tag,
                "output: <#{stdout.read}>",
                "error: <#{stderr.read}>",
              ].join("\n")
            end
            return
          end
          File.open(text.path) do |text_input|
            yield(TextData.new(text_input.read, source_data: data))
          end
        end
      end

      private
      def find_command
        candidates = [
          @options[:abiword],
          ENV["ABIWORD"],
          "abiword",
        ]
        candidates.each do |candidate|
          next if candidate.nil?
          command = ExternalCommand.new(candidate)
          return command if command.exist?
        end
        nil
      end

      def create_tempfiles(data)
        basename = File.basename(data.path)
        text = Tempfile.new([basename, ".txt"])
        stdout = Tempfile.new([basename, ".stdout.log"])
        stderr = Tempfile.new([basename, ".stderr.log"])
        begin
          yield(text, stdout, stderr)
        ensure
          text.close!
          stdout.close!
          stderr.close!
        end
      end

      def log_tag
        "[decomposer][abiword]"
      end
    end
  end
end
