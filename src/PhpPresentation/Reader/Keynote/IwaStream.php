<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Reader\Keynote;

use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;

/**
 * The archives a decompressed IWA component is made of.
 *
 * Each archive opens with the length of a `TSP.ArchiveInfo` message, which names the identifier of
 * the archive and, for every message that follows it, the type of that message and its length. The
 * messages themselves are handed back as bytes : which Apple message a type stands for is the
 * caller's business.
 */
class IwaStream
{
    /**
     * Identifier of the archive : `TSP.ArchiveInfo.identifier`.
     *
     * @var int
     */
    protected const FIELD_IDENTIFIER = 1;

    /**
     * The messages of the archive : `TSP.ArchiveInfo.message_infos`.
     *
     * @var int
     */
    protected const FIELD_MESSAGE_INFOS = 2;

    /**
     * Type of a message : `TSP.MessageInfo.type`.
     *
     * @var int
     */
    protected const FIELD_MESSAGE_TYPE = 1;

    /**
     * Length of a message : `TSP.MessageInfo.length`.
     *
     * @var int
     */
    protected const FIELD_MESSAGE_LENGTH = 3;

    /**
     * Reads every archive of a decompressed component.
     *
     * @param string $data A component, once the Snappy frames are decompressed
     * @param string $context Name of the component, for error messages
     *
     * @return array<int, array{identifier: int, messages: array<int, array{type: int, payload: string}>}>
     */
    public static function parse(string $data, string $context = ''): array
    {
        $archives = [];
        $offset = 0;
        $length = strlen($data);

        while ($offset < $length) {
            $infoLength = Protobuf::readVarint($data, $offset, $context);
            if ($offset + $infoLength > $length) {
                throw new InvalidFileFormatException($context, self::class, 'truncated IWA archive header');
            }
            $info = Protobuf::decode(substr($data, $offset, $infoLength), $context);
            $offset += $infoLength;

            $archives[] = [
                'identifier' => Protobuf::getInt($info, self::FIELD_IDENTIFIER) ?? 0,
                'messages' => self::readMessages($data, $offset, $info, $context),
            ];
        }

        return $archives;
    }

    /**
     * The messages an archive header announces, read in the order it announces them.
     *
     * @param array<int, array<int, int|string>> $info
     *
     * @return array<int, array{type: int, payload: string}>
     */
    protected static function readMessages(string $data, int &$offset, array $info, string $context): array
    {
        $messages = [];

        foreach (Protobuf::getStrings($info, self::FIELD_MESSAGE_INFOS) as $messageInfo) {
            $fields = Protobuf::decode($messageInfo, $context);
            $messageLength = Protobuf::getInt($fields, self::FIELD_MESSAGE_LENGTH) ?? 0;
            if ($offset + $messageLength > strlen($data)) {
                throw new InvalidFileFormatException($context, self::class, 'truncated IWA message');
            }
            $messages[] = [
                'type' => Protobuf::getInt($fields, self::FIELD_MESSAGE_TYPE) ?? 0,
                'payload' => substr($data, $offset, $messageLength),
            ];
            $offset += $messageLength;
        }

        return $messages;
    }
}
