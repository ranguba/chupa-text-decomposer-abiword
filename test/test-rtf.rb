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

class TestRtf < Test::Unit::TestCase
  include FixtureHelper

  def setup
    @decomposer = ChupaText::Decomposers::AbiWord.new({})
  end

  def fixture_path(*components)
    super("rtf", *components)
  end

  sub_test_case("target?") do
    sub_test_case("extension") do
      def create_data(uri)
        data = ChupaText::Data.new
        data.body = ""
        data.uri = uri
        data
      end

      def test_doc
        assert_true(@decomposer.target?(create_data("document.rtf")))
      end
    end

    sub_test_case("mime-type") do
      def create_data(mime_type)
        data = ChupaText::Data.new
        data.mime_type = mime_type
        data
      end

      def test_rich_text_format
        mime_type = "application/rtf"
        assert_true(@decomposer.target?(create_data(mime_type)))
      end
    end
  end

  sub_test_case("decompose") do
    include DecomposeHelper

    sub_test_case("one page") do
      def test_body
        assert_equal(["Page1"], decompose.collect(&:body))
      end

      private
      def decompose
        super(fixture_path("one-page.rtf"))
      end
    end

    sub_test_case("multi pages") do
      def test_body
        assert_equal([<<-BODY.chomp], decompose.collect(&:body))
Page1
Page2
        BODY
      end

      private
      def decompose
        super(fixture_path("multi-pages.rtf"))
      end
    end
  end
end
